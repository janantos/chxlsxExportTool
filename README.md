### Description:
  ClickHouse to XLSX exporter

### Usage:
  chxlsxExportTool [options]

### Options:
  --clickhouse-uri <clickhouse-uri>            ClickHouse URI protocol://hostname:port/database [default: http://localhost:8123/default]
  
  --clickhouse-user <clickhouse-user>          ClickHouse User [default: default]
  
  --clickhouse-password <clickhouse-password>  ClickHouse Password []
  
  --query <query>                              ClickHouse Query []
  
  --output-filename <output-filename>          Output File Name without suffix [default: export]
  
  --split-rows <split-rows>                    Split Excel file every [x] rows [default: 400000]
  
  --datetime-format <datetime-format>          DateTime format [default: dd/mm/yyyy hh:mm:ss]
  
  --version                                    Show version information
  
  -h, --help                               Show help and usage information

### OS X users note:
Application is not signed by Apple 3rd party developer certificate

you may need to run the following (example for arm64 binary) 
```
chmod 755 chxslxExportTool_osx-arm64
spctl --add chxslxExportTool_osx-arm64
open chxslxExportTool_osx-arm64

```
