// RapidTillCsv
// Konverterar Ashton-Tate RapidFile-filer (.RPD) och dBase-filer (.DBF) till
// CSV (UTF-8, semikolon) och XLSX, med korrekta å, ä, ö (DOS-teckentabell 850).
//
// Användning:
//   RapidTillCsv.exe                      alla .RPD/.DBF i aktuell mapp och undermappar
//   RapidTillCsv.exe fil.RPD mapp ...     angivna filer/mappar (dra och släpp går bra)
// Flaggor:
//   --komma        komma som avgränsare i CSV (standard är semikolon)
//   --tab          tabb som avgränsare i CSV
//   --ingen-xlsx   skriv bara CSV
//
// Kompileras mot .NET Framework 4.5: använd bara API:er som finns där
// (t.ex. TrimEnd(null), inte TrimEnd() utan argument).
//
// RPD-formatet (RapidFile, magic E3 92 87 45), utrett från filerna:
//   0x04 u16  storlek på filhuvudet; fältdefinitionerna börjar på värdet + 4
//   0x06 u16  antal poster
//   0x08 u8   antal fält
//   Fältdefinition: u16 längd, [u8 n, formeltext], [u8 n, kompilerad formel],
//                   [u8 n, namn], [u8 n, attribut: typ, fältnummer, bredd, ...]
//   Resten av filen är en kedja av block, var och ett med u16 längd (inklusive
//   de två längdbytena): först vyer, rapporter och formulär, sist posterna.
//   Post: ett [u8 n, text] per fält. Nyare filer (0x0A = 0D 02) utelämnar
//   tomma fält i slutet av posten.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;

namespace RapidTillCsv
{
    sealed class Table
    {
        public List<string> Columns = new List<string>();
        public List<string[]> Rows = new List<string[]>();
        public List<string> Warnings = new List<string>();
    }

    static class Cp850
    {
        const string High =
            "ÇüéâäàåçêëèïîìÄÅ" +
            "ÉæÆôöòûùÿÖÜø£Ø×ƒ" +
            "áíóúñÑªº¿®¬½¼¡«»" +
            "░▒▓│┤ÁÂÀ©╣║╗╝¢¥┐" +
            "└┴┬├─┼ãÃ╚╔╩╦╠═╬¤" +
            "ðÐÊËÈıÍÎÏ┘┌█▄¦Ì▀" +
            "ÓßÔÒõÕµþÞÚÛÙýÝ¯´" +
            "­±‗¾¶§÷¸°¨·¹³²■ ";

        // Avkodar och ersätter styrtecken (CR, LF, tabb m.fl.) med mellanslag.
        public static string Decode(byte[] d, int off, int len)
        {
            var sb = new StringBuilder(len);
            for (int i = off; i < off + len; i++)
            {
                byte b = d[i];
                if (b >= 0x80) sb.Append(High[b - 0x80]);
                else if (b < 0x20 || b == 0x7F) sb.Append(' ');
                else sb.Append((char)b);
            }
            return sb.ToString();
        }
    }

    static class RpdReader
    {
        sealed class Field
        {
            public string Name;
            public int Index;
        }

        static int U16(byte[] d, int p) { return d[p] | (d[p + 1] << 8); }

        public static Table Read(string path)
        {
            byte[] d = File.ReadAllBytes(path);
            if (d.Length < 0x720 || d[0] != 0xE3 || d[1] != 0x92 || d[2] != 0x87 || d[3] != 0x45)
                throw new InvalidDataException("inte en RapidFile-fil (okänt filhuvud)");

            int nrec = U16(d, 6);
            int nf = d[8];
            if (nf == 0) throw new InvalidDataException("filen har inga fält");

            var fields = new List<Field>();
            int p = U16(d, 4) + 4;
            for (int i = 0; i < nf; i++)
            {
                if (p + 2 > d.Length) throw new InvalidDataException("fältdefinitionerna är trasiga");
                int len = U16(d, p);
                if (len < 6 || p + len > d.Length) throw new InvalidDataException("fältdefinition " + (i + 1) + " är trasig");
                int q = p + 2;
                q += 1 + d[q];                 // formel som text
                q += 1 + d[q];                 // kompilerad formel
                int nameLen = d[q];
                string name = Cp850.Decode(d, q + 1, nameLen).Trim();
                q += 1 + nameLen;
                int attrLen = d[q];
                int index = attrLen >= 2 ? d[q + 2] : i;
                fields.Add(new Field { Name = name.Length > 0 ? name : "Fält" + (i + 1), Index = index });
                p += len;
            }

            // Allt efter fältdefinitionerna är en kedja av block med u16-längd:
            // först vyer, rapporter och formulär, sist posterna. Posterna är
            // alltså de sista blocken i kedjan.
            var blocks = new List<int>();
            int pos = p;
            while (pos + 2 <= d.Length)
            {
                int len = U16(d, pos);
                if (len < 2 || pos + len > d.Length) break;
                blocks.Add(pos);
                pos += len;
            }
            if (pos != d.Length)
                throw new InvalidDataException("okänd filstruktur vid position 0x" + pos.ToString("X") +
                                               " (blockkedjan går inte ända till filslut)");

            // Antalet poster lagras i 16 bitar. Har filen fler än 65 535 poster
            // syns det som fler postliknande block i slutet av kedjan.
            int tail = 0;
            for (int i = blocks.Count - 1; i >= 0 && FieldCount(d, blocks[i], nf) >= 0; i--) tail++;
            int count = nrec;
            while (count + 65536 <= tail) count += 65536;
            if (count > tail)
                throw new InvalidDataException("filhuvudet säger " + count + " poster men bara " + tail + " hittades");

            var t = new Table();

            // Kolumnordning efter fältnummer om de är 0..nf-1, annars definitionsordning.
            var byIndex = new Field[nf];
            bool indexOk = true;
            foreach (var f in fields)
            {
                if (f.Index < 0 || f.Index >= nf || byIndex[f.Index] != null) { indexOk = false; break; }
                byIndex[f.Index] = f;
            }
            if (!indexOk) byIndex = fields.ToArray();
            foreach (var f in byIndex) t.Columns.Add(f.Name);

            // En post: ett [u8 n, text] per fält. Tomma fält sist i posten
            // utelämnas i nyare filer och blir tomma kolumner här.
            for (int i = blocks.Count - count; i < blocks.Count; i++)
            {
                int q = blocks[i] + 2;
                int end = blocks[i] + U16(d, blocks[i]);
                var row = new string[nf];
                for (int k = 0; k < nf; k++)
                {
                    if (q >= end) { row[k] = ""; continue; }
                    int l = d[q];
                    row[k] = Cp850.Decode(d, q + 1, l).TrimEnd(null);
                    q += 1 + l;
                }
                t.Rows.Add(row);
            }
            return t;
        }

        // Antal fält i blocket vid q om det ser ut som en post (högst nf fält
        // som går jämnt ut mot blockets slut), annars -1.
        static int FieldCount(byte[] d, int q, int nf)
        {
            int end = q + U16(d, q), t = q + 2, c = 0;
            while (t < end) { t += 1 + d[t]; c++; }
            return (t == end && c <= nf) ? c : -1;
        }
    }

    static class DbfReader
    {
        public static Table Read(string path)
        {
            byte[] d = File.ReadAllBytes(path);
            if (d.Length < 33) throw new InvalidDataException("för kort för att vara en DBF-fil");
            int nrec = BitConverter.ToInt32(d, 4);
            int headerLen = d[8] | (d[9] << 8);
            int recLen = d[10] | (d[11] << 8);
            if (headerLen < 33 || recLen < 1 || headerLen > d.Length)
                throw new InvalidDataException("inte en dBase-fil (okänt filhuvud)");

            var names = new List<string>();
            var types = new List<char>();
            var widths = new List<int>();
            var decs = new List<int>();
            for (int p = 32; p + 32 <= headerLen && d[p] != 0x0D; p += 32)
            {
                int n = 0;
                while (n < 11 && d[p + n] != 0) n++;
                names.Add(Cp850.Decode(d, p, n).Trim());
                types.Add((char)d[p + 11]);
                widths.Add(d[p + 16]);
                decs.Add(d[p + 17]);
            }
            if (names.Count == 0) throw new InvalidDataException("DBF-filen har inga fält");

            int sum = 1;
            foreach (int w in widths) sum += w;
            if (sum != recLen)
            {
                // Clipper/FoxPro: långa teckenfält använder decimalbyten som högbyte.
                int alt = 1;
                for (int i = 0; i < widths.Count; i++)
                    alt += types[i] == 'C' ? widths[i] + decs[i] * 256 : widths[i];
                if (alt == recLen)
                    for (int i = 0; i < widths.Count; i++)
                        if (types[i] == 'C') { widths[i] += decs[i] * 256; decs[i] = 0; }
            }

            var t = new Table();
            t.Columns.AddRange(names);
            int deleted = 0;
            for (int r = 0; r < nrec; r++)
            {
                int off = headerLen + r * recLen;
                if (off + recLen > d.Length)
                {
                    t.Warnings.Add("filen tar slut efter " + r + " av " + nrec + " poster");
                    break;
                }
                if (d[off] == (byte)'*') { deleted++; continue; }
                var row = new string[names.Count];
                int o = off + 1;
                for (int i = 0; i < names.Count; i++)
                {
                    string v = Cp850.Decode(d, o, Math.Min(widths[i], off + recLen - o)).TrimEnd(null);
                    row[i] = Format(types[i], v);
                    o += widths[i];
                }
                t.Rows.Add(row);
            }
            if (deleted > 0) t.Warnings.Add(deleted + " raderade poster hoppades över");
            return t;
        }

        static string Format(char type, string v)
        {
            v = v.Trim();
            if (type == 'N' || type == 'F')
            {
                decimal x;
                if (decimal.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out x))
                {
                    if (x == decimal.Truncate(x)) return decimal.Truncate(x).ToString(CultureInfo.InvariantCulture);
                    return x.ToString(CultureInfo.InvariantCulture).Replace('.', ',');
                }
                return v;
            }
            if (type == 'D' && v.Length == 8)
                return v.Substring(0, 4) + "-" + v.Substring(4, 2) + "-" + v.Substring(6, 2);
            return v;
        }
    }

    static class CsvWriter
    {
        public static void Write(Table t, string path, char sep)
        {
            using (var w = new StreamWriter(path, false, new UTF8Encoding(true)))
            {
                w.NewLine = "\r\n";
                WriteRow(w, t.Columns.ToArray(), sep);
                foreach (var row in t.Rows) WriteRow(w, row, sep);
            }
        }

        static void WriteRow(TextWriter w, string[] values, char sep)
        {
            for (int i = 0; i < values.Length; i++)
            {
                if (i > 0) w.Write(sep);
                string v = values[i] ?? "";
                bool quote = v.IndexOf(sep) >= 0 || v.IndexOf('"') >= 0 || v.IndexOf('\r') >= 0 ||
                             v.IndexOf('\n') >= 0 || (v.Length > 0 && (v[0] == ' ' || v[v.Length - 1] == ' '));
                if (quote) w.Write("\"" + v.Replace("\"", "\"\"") + "\"");
                else w.Write(v);
            }
            w.WriteLine();
        }
    }

    // Minimal XLSX utan externa beroenden. Kolumner där alla värden är heltal
    // (utan inledande nollor) blir tal, allt annat text, så att t.ex. "001",
    // "3/8" och "31:17" inte tolkas om av Excel.
    static class XlsxWriter
    {
        const string NsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const string NsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string NsPkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";
        const string Head = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";

        public static void Write(Table t, string path, string sheetName)
        {
            int nc = t.Columns.Count;
            var numeric = new bool[nc];
            var width = new int[nc];
            for (int c = 0; c < nc; c++)
            {
                bool any = false, all = true;
                int max = t.Columns[c].Length;
                foreach (var row in t.Rows)
                {
                    string v = row[c];
                    if (string.IsNullOrEmpty(v)) continue;
                    any = true;
                    if (all && !IsPlainInteger(v)) all = false;
                    if (v.Length > max) max = v.Length;
                }
                numeric[c] = any && all;
                width[c] = Math.Max(6, Math.Min(60, max)) + 2;
            }

            string lastRef = Col(nc - 1) + (t.Rows.Count + 1);
            var sb = new StringBuilder(64 * 1024);
            sb.Append(Head);
            sb.Append("<worksheet xmlns=\"").Append(NsMain).Append("\" xmlns:r=\"").Append(NsRel).Append("\">");
            sb.Append("<dimension ref=\"A1:").Append(lastRef).Append("\"/>");
            sb.Append("<sheetViews><sheetView workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/></sheetView></sheetViews>");
            sb.Append("<sheetFormatPr defaultRowHeight=\"15\"/><cols>");
            for (int c = 0; c < nc; c++)
                sb.Append("<col min=\"").Append(c + 1).Append("\" max=\"").Append(c + 1)
                  .Append("\" width=\"").Append(width[c]).Append("\" customWidth=\"1\"/>");
            sb.Append("</cols><sheetData>");

            sb.Append("<row r=\"1\">");
            for (int c = 0; c < nc; c++) TextCell(sb, Col(c) + "1", t.Columns[c], 1);
            sb.Append("</row>");
            for (int r = 0; r < t.Rows.Count; r++)
            {
                string rn = (r + 2).ToString(CultureInfo.InvariantCulture);
                sb.Append("<row r=\"").Append(rn).Append("\">");
                var row = t.Rows[r];
                for (int c = 0; c < nc; c++)
                {
                    string v = row[c];
                    if (string.IsNullOrEmpty(v)) continue;
                    if (numeric[c]) sb.Append("<c r=\"").Append(Col(c)).Append(rn).Append("\"><v>").Append(v).Append("</v></c>");
                    else TextCell(sb, Col(c) + rn, v, 2);
                }
                sb.Append("</row>");
            }
            sb.Append("</sheetData><autoFilter ref=\"A1:").Append(lastRef).Append("\"/></worksheet>");

            string styles = Head +
                "<styleSheet xmlns=\"" + NsMain + "\">" +
                "<fonts count=\"2\"><font><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/></font>" +
                "<font><b/><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/></font></fonts>" +
                "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>" +
                "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>" +
                "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
                "<cellXfs count=\"3\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>" +
                "<xf numFmtId=\"49\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyNumberFormat=\"1\"/>" +
                "<xf numFmtId=\"49\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/></cellXfs>" +
                "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>" +
                "</styleSheet>";

            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                Add(zip, "[Content_Types].xml", Head +
                    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                    "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                    "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                    "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                    "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                    "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" +
                    "</Types>");
                Add(zip, "_rels/.rels", Head +
                    "<Relationships xmlns=\"" + NsPkgRel + "\">" +
                    "<Relationship Id=\"rId1\" Type=\"" + NsRel + "/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                    "</Relationships>");
                Add(zip, "xl/workbook.xml", Head +
                    "<workbook xmlns=\"" + NsMain + "\" xmlns:r=\"" + NsRel + "\">" +
                    "<sheets><sheet name=\"" + Esc(sheetName) + "\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
                Add(zip, "xl/_rels/workbook.xml.rels", Head +
                    "<Relationships xmlns=\"" + NsPkgRel + "\">" +
                    "<Relationship Id=\"rId1\" Type=\"" + NsRel + "/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
                    "<Relationship Id=\"rId2\" Type=\"" + NsRel + "/styles\" Target=\"styles.xml\"/>" +
                    "</Relationships>");
                Add(zip, "xl/styles.xml", styles);
                Add(zip, "xl/worksheets/sheet1.xml", sb.ToString());
            }
        }

        public static string SheetName(string fileName)
        {
            var sb = new StringBuilder();
            foreach (char ch in fileName)
                sb.Append("[]:*?/\\".IndexOf(ch) >= 0 ? '_' : ch);
            string s = sb.ToString().Trim('\'', ' ');
            if (s.Length > 31) s = s.Substring(0, 31);
            return s.Length == 0 ? "Data" : s;
        }

        static bool IsPlainInteger(string v)
        {
            int i = v[0] == '-' ? 1 : 0;
            int digits = v.Length - i;
            if (digits < 1 || digits > 15) return false;
            if (v[i] == '0' && digits > 1) return false;
            for (; i < v.Length; i++) if (v[i] < '0' || v[i] > '9') return false;
            return v != "-0";
        }

        static void TextCell(StringBuilder sb, string cellRef, string v, int style)
        {
            sb.Append("<c r=\"").Append(cellRef).Append("\" t=\"inlineStr\" s=\"").Append(style)
              .Append("\"><is><t xml:space=\"preserve\">").Append(Esc(v)).Append("</t></is></c>");
        }

        static string Col(int c)
        {
            string s = "";
            for (c++; c > 0; c = (c - 1) / 26) s = (char)('A' + (c - 1) % 26) + s;
            return s;
        }

        static string Esc(string v)
        {
            var sb = new StringBuilder(v.Length + 8);
            foreach (char ch in v)
            {
                switch (ch)
                {
                    case '&': sb.Append("&amp;"); break;
                    case '<': sb.Append("&lt;"); break;
                    case '>': sb.Append("&gt;"); break;
                    case '"': sb.Append("&quot;"); break;
                    default: sb.Append(ch < 0x20 ? ' ' : ch); break;
                }
            }
            return sb.ToString();
        }

        static void Add(ZipArchive zip, string name, string content)
        {
            var e = zip.CreateEntry(name, CompressionLevel.Optimal);
            using (var s = e.Open())
            {
                byte[] b = new UTF8Encoding(false).GetBytes(content);
                s.Write(b, 0, b.Length);
            }
        }
    }

    static class Program
    {
        [DllImport("kernel32.dll")]
        static extern uint GetConsoleProcessList(uint[] list, uint count);

        // Sant när programmet startats med dubbelklick eller dra-och-släpp,
        // så att fönstret inte stängs innan man hunnit läsa.
        static bool OwnsConsoleWindow()
        {
            try { return GetConsoleProcessList(new uint[2], 2) == 1; }
            catch { return false; }
        }

        static int Main(string[] args)
        {
            int rc = Run(args);
            if (OwnsConsoleWindow())
            {
                Console.WriteLine();
                Console.Write("Tryck på valfri tangent för att stänga...");
                try { Console.ReadKey(true); } catch { }
            }
            return rc;
        }

        static int Run(string[] args)
        {
            char sep = ';';
            bool xlsx = true;
            var inputs = new List<string>();
            foreach (string a in args)
            {
                string k = a.ToLowerInvariant();
                if (k == "--komma") sep = ',';
                else if (k == "--tab") sep = '\t';
                else if (k == "--semikolon") sep = ';';
                else if (k == "--ingen-xlsx") xlsx = false;
                else if (k == "-h" || k == "--help" || k == "/?") { Usage(); return 0; }
                else if (k.StartsWith("--")) { Console.WriteLine("Okänd flagga: " + a); Usage(); return 2; }
                else inputs.Add(a);
            }

            var files = new List<string>();
            if (inputs.Count == 0)
            {
                string dir = Directory.GetCurrentDirectory();
                Console.WriteLine("Söker .RPD- och .DBF-filer i " + dir);
                Collect(dir, files);
            }
            foreach (string inp in inputs)
            {
                if (Directory.Exists(inp)) Collect(Path.GetFullPath(inp), files);
                else if (File.Exists(inp)) files.Add(Path.GetFullPath(inp));
                else Console.WriteLine("Hittar inte: " + inp);
            }
            files.Sort(StringComparer.OrdinalIgnoreCase);
            for (int i = files.Count - 1; i > 0; i--)
                if (string.Equals(files[i], files[i - 1], StringComparison.OrdinalIgnoreCase)) files.RemoveAt(i);
            if (files.Count == 0)
            {
                Console.WriteLine("Inga .RPD- eller .DBF-filer att konvertera.");
                return 1;
            }

            int ok = 0, failed = 0;
            foreach (string f in files)
            {
                string ext = Path.GetExtension(f).ToUpperInvariant();
                string dir = Path.GetDirectoryName(f);
                string stem = Path.GetFileNameWithoutExtension(f);
                string outStem = ext == ".DBF" ? stem + "_dbf" : stem;
                try
                {
                    Table t = ext == ".DBF" ? DbfReader.Read(f) : RpdReader.Read(f);
                    string csv = Path.Combine(dir, outStem + ".csv");
                    CsvWriter.Write(t, csv, sep);
                    string made = Path.GetFileName(csv);
                    if (xlsx)
                    {
                        string x = Path.Combine(dir, outStem + ".xlsx");
                        XlsxWriter.Write(t, x, XlsxWriter.SheetName(stem));
                        made += " + " + Path.GetFileName(x);
                    }
                    Console.WriteLine("OK   " + Path.GetFileName(f) + " -> " + made +
                                      "  (" + t.Rows.Count + " poster, " + t.Columns.Count + " fält)");
                    foreach (string w in t.Warnings) Console.WriteLine("     obs: " + w);
                    ok++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("FEL  " + Path.GetFileName(f) + ": " + ex.Message);
                    failed++;
                }
            }
            Console.WriteLine();
            Console.WriteLine(ok + " fil(er) konverterade" + (failed > 0 ? ", " + failed + " misslyckades." : "."));
            return failed > 0 ? 1 : 0;
        }

        static void Collect(string dir, List<string> files)
        {
            try
            {
                foreach (string f in Directory.GetFiles(dir))
                {
                    string e = Path.GetExtension(f).ToUpperInvariant();
                    if (e == ".RPD" || e == ".DBF") files.Add(f);
                }
                foreach (string sub in Directory.GetDirectories(dir)) Collect(sub, files);
            }
            catch (UnauthorizedAccessException) { }
            catch (IOException) { }
        }

        static void Usage()
        {
            Console.WriteLine("RapidTillCsv - konverterar RapidFile (.RPD) och dBase (.DBF) till CSV och Excel");
            Console.WriteLine();
            Console.WriteLine("  RapidTillCsv.exe                 alla .RPD/.DBF i aktuell mapp och undermappar");
            Console.WriteLine("  RapidTillCsv.exe fil.RPD mapp    angivna filer/mappar (dra och släpp går bra)");
            Console.WriteLine();
            Console.WriteLine("  --komma       komma som avgränsare (standard: semikolon, som svensk Excel vill ha)");
            Console.WriteLine("  --tab         tabb som avgränsare");
            Console.WriteLine("  --ingen-xlsx  skriv bara CSV");
            Console.WriteLine();
            Console.WriteLine("NAMN.RPD ger NAMN.csv och NAMN.xlsx, NAMN.DBF ger NAMN_dbf.csv och NAMN_dbf.xlsx.");
        }
    }
}
