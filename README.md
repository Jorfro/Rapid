# RapidTillCsv

Konverterar Ashton-Tates **RapidFile** (`.RPD`, DOS, ca 1986–1990) och **dBase** (`.DBF`) till CSV och Excel (`.xlsx`), med rätt å, ä och ö. Ingen DOSBox behövs.

## Användning

- **Dubbelklicka** på `RapidTillCsv.exe` så konverteras alla `.RPD`/`.DBF` i mappen den startas i, inklusive undermappar.
- **Dra och släpp** en eller flera filer eller mappar på exe-filen.
- **Kommandorad:** `RapidTillCsv.exe P:\Rapid\Filer\TOMPOWSK.RPD`
- **Hel disk utan att röra originalen:** `RapidTillCsv.exe E:\ --ut D:\Export`

| Flagga | Betydelse |
|---|---|
| `--ut MAPP` | skriv resultatet i MAPP, med samma undermappar som i källan, i stället för bredvid källfilerna |
| `--komma` | komma som avgränsare i CSV (standard är semikolon, som svensk Excel vill ha) |
| `--tab` | tabb som avgränsare |
| `--ingen-xlsx` | skriv bara CSV |

Utfilerna får samma namn som källfilen:

- `NAMN.RPD` → `NAMN.csv` + `NAMN.xlsx`
- `NAMN.DBF` → `NAMN_dbf.csv` + `NAMN_dbf.xlsx`

Utan `--ut` hamnar de bredvid källfilen, och en befintlig `NAMN.csv`/`NAMN.xlsx` där skrivs över. Källfilerna ändras aldrig. Om två källfiler i samma körning skulle ge samma utfil (t.ex. två mappar med `TOMPOWSK.RPD` och samma `--ut`) får den andra suffixet `_2`.

### Loggfil

Varje körning skriver `RapidTillCsv_logg_<datum>_<tid>.txt` i utmappen, eller i källmappen om `--ut` saknas. Den listar varje fil med `OK` eller `FEL` och eventuella anmärkningar (`obs`). Sist kommer en sammanfattning med alla filer som gav FEL och alla mappar som inte gick att läsa.

En fil med fel stoppar inte körningen. Anmärkningar för DBF-filer:

- **Windows-teckentabell:** filhuvudet anger Windows (1252), och filen avkodas därefter.
- **Memofält:** texten ligger i en separat `.DBT`- eller `.FPT`-fil och följer inte med.
- **Fält som inte är text:** binära FoxPro-fält som `I`, `Y` och `T` kan bli skräp.

### CSV eller XLSX?

CSV-filen är UTF-8 med BOM, semikolon och CRLF, så den öppnas med rätt tecken vid dubbelklick i svensk Excel. Excel gissar dock datatyp när den öppnar CSV: `001` blir `1`, `3/8` blir ett datum, `31:17` blir ett klockslag och `+` eller `---sås` tolkas som formler. Sådana värden finns i filerna.

Öppna därför **`.xlsx`** i Excel. Där sparas alla textfält som text, och bara kolumner som enbart innehåller heltal utan inledande nollor blir tal. Översta raden är låst och har filter.

## Bygga

`dist/RapidTillCsv.exe` kräver bara .NET Framework 4.5 eller senare, som redan finns i Windows 10 och 11. Koden får därför bara använda API:er som finns i .NET Framework, t.ex. `TrimEnd(null)` och inte `TrimEnd()` utan argument.

Bygga själv med .NET SDK:

```
dotnet build src -c Release
dotnet publish src -c Release -r win-x64 --self-contained -p:PublishSingleFile=true
```

Exe-filen i `dist/` är byggd med Mono (`mcs -sdk:4.5 -r:System.IO.Compression.dll src/Program.cs`) för att köras på .NET Framework utan installation.

## RPD-formatet

Det finns ingen publicerad beskrivning. Detta är utrett från filerna här:

| Offset | Typ | Innehåll |
|---|---|---|
| `0x00` | 4 byte | magic `E3 92 87 45` |
| `0x04` | u16 | storlek på filhuvudet (`0x718`). Fältdefinitionerna börjar på värdet + 4 |
| `0x06` | u16 | antal poster |
| `0x08` | u8 | antal fält |

- **Fältdefinition:** u16 längd, sedan fyra längdprefixade delar: formel som text, kompilerad formel, namn (13 tecken, utfyllt med blanksteg) och attribut (typ, fältnummer, bredd …). Typerna är `A` (text), `N` (tal), `a` och `0` (beräknade fält). Beräknade värden är sparade i posterna.
- Resten av filen är en **kedja av block**, vart och ett med u16 längd (inklusive de två längdbytena). Först kommer vyer, rapporter, formulär och sorteringsordningar, och sist posterna. Kedjan går ända till filslut, så posterna är de sista *antal poster* blocken.
- **Post:** ett `[u8 längd][text]` per fält i fältnummerordning. Allt är text, även talfält (därför kan ett talfält innehålla t.ex. `933B`). Nyare filer (byte `0x0A`–`0x0B` = `0D 02`, t.ex. Norhalla och Sydhalla) utelämnar tomma fält i slutet av posten.
- Antal poster lagras i 16 bitar. Har en fil fler än 65 535 poster märks det som fler postliknande block i slutet, och då räknas de med.
- **Teckentabell:** DOS 850. Å/Ä/Ö är `8F/8E/99` och å/ä/ö `86/84/94`. Tabell 850 behövs i stället för 437 eftersom `EF` används som apostrof (´).

## Kontroll

Utdata från RPD har jämförts cell för cell med motsvarande DBF-exporter:

| Fil | Resultat |
|---|---|
| `1732-36` | alla 23 320 celler lika |
| `1737-46` | lika, förutom att DBF har tappat inledande nollor (`0389` → `389`, `001` → `1`) |
| `TOMPOWSK` | lika, förutom 44 celler där DBF-exporten kapat text eller gjort om text i talfält till `0` (t.ex. `se Sählberg`). RPD-versionen är den kompletta. |
| `ORBY1900` | DBF-filen är exporterad nio dagar senare, med nya fält och annan sortering. Far/Mor stämmer för 4 591 av 4 592 personer. |

Resultatet är detsamma oavsett om programmet körs på .NET Framework (Mono) eller .NET 8.
