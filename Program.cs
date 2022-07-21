using System;
using System.Text.Json;
using System.CommandLine;
using ClickHouse.Client.ADO;
using ClickHouse.Client.Utility;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

[assembly: System.Reflection.AssemblyVersion("1.0.1.*")]



Option<String> chURIOption = new(
  name: "--clickhouse-uri",
  description: "ClickHouse URI protocol://hostname:port/database",
  getDefaultValue: () => "http://localhost:8123/default"
    );
Option<String?> chQueryOption = new(
  name: "--query",
  description: "ClickHouse Query",
  getDefaultValue: () => null
    );
Option<String> chUserOption = new(
  name: "--clickhouse-user",
  description: "ClickHouse User",
  getDefaultValue: () => "default"
    );
Option<String?> chPasswordOption = new(
  name: "--clickhouse-password",
  description: "ClickHouse Password",
  getDefaultValue: () => null
    );
Option<String> outputFileNameOption = new(
  name: "--output-filename",
  description: "Output File Name without suffix",
  getDefaultValue: () => "export"
    );
Option<Int32> splitRowsOption = new(
  name: "--split-rows",
  description: "Split Excel file every [x] rows",
  getDefaultValue: () => 400000
    );
Option<String> datetimeFormatOption = new(
  name: "--datetime-format",
  description: "DateTime format",
  getDefaultValue: () => "dd/mm/yyyy hh:mm:ss"
    );

RootCommand rootCommand = new(description: "ClickHouse to XLSX exporter"){
  chURIOption,
  chUserOption,
  chPasswordOption,
  chQueryOption,
  outputFileNameOption,
  splitRowsOption,
  datetimeFormatOption
};



rootCommand.SetHandler(
  async (String chURI, String? chQuery, String chUser, String? chPassword, String outputFileName, Int32 maxrows, String datetimeFormat) =>
  {
      
      if (chQuery is null)
      {
          TextWriter errorWriter = Console.Error;
          errorWriter.WriteLine("--query parameter not defined");
      }
      else
      {
          Uri uri = new Uri(chURI);
          String chHost = uri.Host;
          ushort chPort = (ushort)uri.Port;
          String chProtocol = uri.Scheme;
          String chDatabase = uri.PathAndQuery.Replace("/","");
          List<CellValues> colTypes = new List<CellValues>();
          var sb = new ClickHouseConnectionStringBuilder();
          sb.Host = chHost;
          sb.Port = chPort;
          sb.Username = chUser;
          sb.Protocol = chProtocol;
          sb.Database = chDatabase;
          sb.Timeout = TimeSpan.FromMinutes(120);
          if (chPassword is not null)
          {
              sb.Password = chPassword;
          }
          using var conn = new ClickHouseConnection();
          conn.ConnectionString = sb.ConnectionString;
          await conn.OpenAsync();
          var cmd = conn.CreateCommand();
          cmd.CommandText = chQuery;
          var rd = await cmd.ExecuteReaderAsync();
          var columnNames = rd.GetColumnNames();

          Console.WriteLine("\nWriting File: " + outputFileName + ".xlsx");
          var spreadsheetDocument = SpreadsheetDocument.Create(outputFileName + ".xlsx", SpreadsheetDocumentType.Workbook);
          WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
          workbookpart.Workbook = new Workbook();
          WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
          worksheetPart.Worksheet = new Worksheet(new SheetData());
          Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
          Sheet exportSheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "export" };

          //-----
          var styleSheet = new Stylesheet();
          NumberingFormats nfs = new NumberingFormats();
          NumberingFormat nf;
          nf = new NumberingFormat();
          nf.NumberFormatId = 165;
          nf.FormatCode = datetimeFormat;
          nfs.Append(nf);
          styleSheet.Append(nfs);

          var Fonts = new Fonts();
          Fonts.Append(new Font()
          {
              FontName = new FontName() { Val = "Calibri" },
              FontSize = new FontSize() { Val = 11 },
              FontFamilyNumbering = new FontFamilyNumbering() { Val = 2 },
          });
          Fonts.Count = (uint)Fonts.ChildElements.Count;
          var Fills = new Fills();
          Fills.Append(new Fill()
          {
              PatternFill = new PatternFill() { PatternType = PatternValues.None }
          });
          Fills.Append(new Fill()
          {
              PatternFill = new PatternFill() { PatternType = PatternValues.Gray125 }
          });
          Fills.Count = (uint)Fills.ChildElements.Count;
          var Borders = new Borders();
          Borders.Append(new Border()
          {
              LeftBorder = new LeftBorder(),
              RightBorder = new RightBorder(),
              TopBorder = new TopBorder(),
              BottomBorder = new BottomBorder(),
              DiagonalBorder = new DiagonalBorder()
          });
          Borders.Count = (uint)Borders.ChildElements.Count;
          var CellStyleFormats = new CellStyleFormats();
          CellStyleFormats.Append(new CellFormat()
          {
              NumberFormatId = 0,
              FontId = 0,
              FillId = 0,
              BorderId = 0
          });
          CellStyleFormats.Count = (uint)CellStyleFormats.ChildElements.Count;
          var CellFormats = new CellFormats();
          CellFormats.Append(new CellFormat()
          {
              BorderId = 0,
              FillId = 0,
              FontId = 0,
              NumberFormatId = 0,
              FormatId = 0,
              ApplyNumberFormat = true
          });
          CellFormats.Append(new CellFormat()
          {
              BorderId = 0,
              FillId = 0,
              FontId = 0,
              NumberFormatId = 165,
              FormatId = 0,
              ApplyNumberFormat = true
          });
          CellFormats.Count = (uint)CellFormats.ChildElements.Count;
          var CellStyles = new CellStyles();
          CellStyles.Append(new CellStyle()
          {
              Name = "Normal",
              FormatId = 0,
              BuiltinId = 0
          });
          CellStyles.Count = (uint)CellStyles.ChildElements.Count;
          styleSheet.Append(Fonts);
          styleSheet.Append(Fills);
          styleSheet.Append(Borders);
          styleSheet.Append(CellStyleFormats);
          styleSheet.Append(CellFormats);
          styleSheet.Append(CellStyles);


          var WorkbookStylesPart = workbookpart.AddNewPart<WorkbookStylesPart>();
          WorkbookStylesPart.Stylesheet = styleSheet;
          WorkbookStylesPart.Stylesheet.Save();
          // -----
          sheets.Append(exportSheet);

          SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());
          Row header = new Row();
          foreach (string colName in columnNames)
          {
              Cell cell = new Cell();
              cell.CellValue = new CellValue(colName);
              cell.DataType = CellValues.String;


              header.Append(cell);
          }
          sheetData.Append(header);


          Int32 rowNum = 0;
          if (rd.HasRows)
          {
              while (rd.Read())
              {
                  rowNum++;
                  if (rowNum % maxrows == 0)
                  {
                      Console.WriteLine("\nWriting File: " + outputFileName+"_"+(rowNum/maxrows).ToString() + ".xlsx");
                      spreadsheetDocument = SpreadsheetDocument.Create(outputFileName+"_"+(rowNum/101).ToString() + ".xlsx", SpreadsheetDocumentType.Workbook);
                      workbookpart = spreadsheetDocument.AddWorkbookPart();
                      workbookpart.Workbook = new Workbook();
                      worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                      worksheetPart.Worksheet = new Worksheet(new SheetData());
                      sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                      exportSheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "export" };

                      //-----
                      styleSheet = new Stylesheet();
                      nfs = new NumberingFormats();
                      //NumberingFormat nf;
                      nf = new NumberingFormat();
                      nf.NumberFormatId = 165;
                      nf.FormatCode = datetimeFormat;
                      nfs.Append(nf);
                      styleSheet.Append(nfs);

                      Fonts = new Fonts();
                      Fonts.Append(new Font()
                      {
                          FontName = new FontName() { Val = "Calibri" },
                          FontSize = new FontSize() { Val = 11 },
                          FontFamilyNumbering = new FontFamilyNumbering() { Val = 2 },
                      });
                      Fonts.Count = (uint)Fonts.ChildElements.Count;
                      Fills = new Fills();
                      Fills.Append(new Fill()
                      {
                          PatternFill = new PatternFill() { PatternType = PatternValues.None }
                      });
                      Fills.Append(new Fill()
                      {
                          PatternFill = new PatternFill() { PatternType = PatternValues.Gray125 }
                      });
                      Fills.Count = (uint)Fills.ChildElements.Count;
                      Borders = new Borders();
                      Borders.Append(new Border()
                      {
                          LeftBorder = new LeftBorder(),
                          RightBorder = new RightBorder(),
                          TopBorder = new TopBorder(),
                          BottomBorder = new BottomBorder(),
                          DiagonalBorder = new DiagonalBorder()
                      });
                      Borders.Count = (uint)Borders.ChildElements.Count;
                      CellStyleFormats = new CellStyleFormats();
                      CellStyleFormats.Append(new CellFormat()
                      {
                          NumberFormatId = 0,
                          FontId = 0,
                          FillId = 0,
                          BorderId = 0
                      });
                      CellStyleFormats.Count = (uint)CellStyleFormats.ChildElements.Count;
                      CellFormats = new CellFormats();
                      CellFormats.Append(new CellFormat()
                      {
                          BorderId = 0,
                          FillId = 0,
                          FontId = 0,
                          NumberFormatId = 0,
                          FormatId = 0,
                          ApplyNumberFormat = true
                      });
                      CellFormats.Append(new CellFormat()
                      {
                          BorderId = 0,
                          FillId = 0,
                          FontId = 0,
                          NumberFormatId = 165,
                          FormatId = 0,
                          ApplyNumberFormat = true
                      });
                      CellFormats.Count = (uint)CellFormats.ChildElements.Count;
                      CellStyles = new CellStyles();
                      CellStyles.Append(new CellStyle()
                      {
                          Name = "Normal",
                          FormatId = 0,
                          BuiltinId = 0
                      });
                      CellStyles.Count = (uint)CellStyles.ChildElements.Count;
                      styleSheet.Append(Fonts);
                      styleSheet.Append(Fills);
                      styleSheet.Append(Borders);
                      styleSheet.Append(CellStyleFormats);
                      styleSheet.Append(CellFormats);
                      styleSheet.Append(CellStyles);


                      WorkbookStylesPart = workbookpart.AddNewPart<WorkbookStylesPart>();
                      WorkbookStylesPart.Stylesheet = styleSheet;
                      WorkbookStylesPart.Stylesheet.Save();
                      // -----
                      sheets.Append(exportSheet);

                      sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());
                      header = new Row();
                      foreach (string colName in columnNames)
                      {
                          Cell cell = new Cell();
                          cell.CellValue = new CellValue(colName);
                          cell.DataType = CellValues.String;


                          header.Append(cell);
                      }
                      sheetData.Append(header);

                  }
                  Row row = new Row();
                  for (int i = 0; i < rd.FieldCount; i++)
                  {

                      Cell cell = new Cell();
                      if (rd[i].GetType() == typeof(System.Byte))
                      {
                          cell.CellValue = new CellValue(rd[i].ToString());
                          cell.DataType = CellValues.Number;
                      }
                      else if (rd[i].GetType() == typeof(System.SByte))
                      {
                          cell.CellValue = new CellValue(rd[i].ToString());
                          cell.DataType = CellValues.Number;
                      }
                      else if (rd[i].GetType() == typeof(System.UInt16))
                      {
                          cell.CellValue = new CellValue(rd[i].ToString());
                          cell.DataType = CellValues.Number;
                      }
                      else if (rd[i].GetType() == typeof(System.UInt32))
                      {
                          cell.CellValue = new CellValue(rd[i].ToString());
                          cell.DataType = CellValues.Number;
                      }
                      else if (rd[i].GetType() == typeof(System.UInt64))
                      {
                          cell.CellValue = new CellValue(rd[i].ToString());
                          cell.DataType = CellValues.Number;
                      }
                      else if (rd[i].GetType() == typeof(System.Int16))
                      {
                          cell.CellValue = new CellValue(rd[i].ToString());
                          cell.DataType = CellValues.Number;
                      }
                      else if (rd[i].GetType() == typeof(System.Int32))
                      {
                          cell.CellValue = new CellValue(rd[i].ToString());
                          cell.DataType = CellValues.Number;
                      }
                      else if (rd[i].GetType() == typeof(System.Int64))
                      {
                          cell.CellValue = new CellValue(rd[i].ToString());
                          cell.DataType = CellValues.Number;
                      }
                      else if (rd[i].GetType() == typeof(System.Decimal))
                      {
                          cell.CellValue = new CellValue(rd[i].ToString());
                          cell.DataType = CellValues.Number;
                      }
                      else if (rd[i].GetType() == typeof(System.Double))
                      {
                          cell.CellValue = new CellValue(rd[i].ToString());
                          cell.DataType = CellValues.Number;
                      }
                      else if (rd[i].GetType() == typeof(System.Single))
                      {
                          cell.CellValue = new CellValue(rd[i].ToString());
                          cell.DataType = CellValues.Number;
                      }
                      else if (rd[i].GetType() == typeof(System.DateTime))
                      {
                          cell.CellValue = new CellValue((DateTime)rd[i]);
                          cell.DataType = CellValues.Date;
                          cell.StyleIndex = 1;
                      }
                      else if (rd[i].GetType() == typeof(System.Boolean))
                      {
                          cell.CellValue = new CellValue(rd[i].ToString());
                          cell.DataType = CellValues.Boolean;
                      } else if (rd[i].GetType() == typeof(System.String))
                      {
                          cell.CellValue = new CellValue(rd[i].ToString());
                          cell.DataType = CellValues.String;
                      } else if (rd[i].GetType() == typeof(System.DBNull))
                      {
                          cell.CellValue = new CellValue(null);
                          cell.DataType = CellValues.String;
                      }else
                      {
                          cell.CellValue = new CellValue(JsonSerializer.Serialize(rd[i]).ToString());   
                          cell.DataType = CellValues.String;
                      }
                      row.Append(cell);
                  }
                  sheetData.Append(row);
                  if (rowNum % 1000 == 0){Console.Write("..."+rowNum.ToString());}
                  if (rowNum % maxrows == maxrows-1)
                  {
                      worksheetPart.Worksheet.Save();
                      workbookpart.Workbook.Save();
                      spreadsheetDocument.Close();

                  }
              }
          }

          worksheetPart.Worksheet.Save();
          workbookpart.Workbook.Save();
          spreadsheetDocument.Close();
          Console.WriteLine("...done");

      }
  },
  chURIOption,
  chQueryOption,
  chUserOption,
  chPasswordOption,
  outputFileNameOption,
  splitRowsOption,
  datetimeFormatOption
);

await rootCommand.InvokeAsync(args);



