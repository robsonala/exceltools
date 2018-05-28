# Exceltools v0.1

C# Core project to convert:
  - CSV to Excel(xlsx)
  - Excel(xlsx) to CSV *Not implemented yet*

### Dependecies
- CsvHelper
- OpenXML SDK
- Newtonsoft.Json

### How to use

**CSV 2 Excel**

```sh
$ ./exceltools excel2csv {INFILE} {OUTFILE} "{JSONSETTINGS}"
```

*JSONSETTINGS*
```
[
    {
        "index": 1, // Column Position
        "width": 10, // Column Size (<= 0 for not set)
        "type": 0 // Format type
    }
    ...
]
```

| FormatID   | FormatDesc               |
|------------|--------------------------|
| 0          | General                  |
| 1          | HEADER (Gray background) |
| 2          | 0.00                     |
| 3          | #,##0                    |
| 4          | #,##0.00                 |
| 5          | 0%                       |
| 6          | 0.00%                    |
| 7          | dd/mm/yyyy               |

*RETURN*
ok -> File generated
error -> File not generated

*Example usage*
```sh
$ ./exceltools excel2csv /opt/example.csv /opt/out/example.xlsx "[{'index':1, 'width': 10, 'type':7}]"
```

**Excel 2 CSV**
** Not implemented yet **
```sh
$ ./exceltools csv2excel {INFILE} {OUTFILE}
```

### Todos
-

License
----

MIT
