using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using TG_Reporting.Model;

namespace TG_Reporting
{
    public class ExcelReportGenerater
    {
        public string GenerateExcelReport(string json)
        {
            TG_Hotel hotelObj = DeserializeJson(json);

            return GenerateExcel(hotelObj);
        }

        public string GenerateExcelReport(Stream stream)
        {
            using (var streamReader = new StreamReader(stream))
            {
                var fileData = streamReader.ReadToEnd();

                TG_Hotel hotelObj = DeserializeJson(fileData);
                return GenerateExcel(hotelObj);
            }
        }

        public string GenerateExcelReport(JsonDocument json)
        {
            TG_Hotel hotelObj = DeserializeJson(json.RootElement.GetRawText());
            return GenerateExcel(hotelObj);
        }

        public string GenerateExcelReport(FileInfo fileinfo)
        {
            var fileData = File.ReadAllText(fileinfo.FullName);
            TG_Hotel hotelObj = DeserializeJson(fileData);
            return GenerateExcel(hotelObj);
        }

        private TG_Hotel DeserializeJson(string fileData)
        {
            return JsonSerializer.Deserialize<TG_Hotel>(fileData, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });
        }

        /// <summary>
        /// Returns the output excel Filename including location
        /// </summary>
        /// <param name="hotelObj"></param>
        /// <returns></returns>
        private string GenerateExcel(TG_Hotel hotelObj)
        {
            Console.WriteLine("Converting the data to desired output model");

            // Need column count and record counts at various places for getting cell references
            colCount = typeof(ReportModel).GetProperties().Length;

            var reportModels = new List<ReportModel>();
            foreach (var hotelRate in hotelObj.HotelRates)
            {
                reportModels.Add(new ReportModel
                {
                    Arrival_Date = hotelRate.TargetDay.Date,
                    Departure_Date = hotelRate.TargetDay.AddDays(hotelRate.Los).Date,
                    Price = hotelRate.Price.NumericFloat,
                    Currency = hotelRate.Price.Currency,
                    RateName = hotelRate.RateName,
                    Adults = hotelRate.Adults,
                    Breakfast_Included = GetBreakfastInclusionValue(hotelRate)
                });
            }

            rowCount = reportModels.Count;

            string fileName = $"{hotelObj.Hotel.HotelID}_Report_{DateTime.Now.ToString("ddMMhhmmss")}.xlsx";
            string outputFolder = Path.Combine(AppContext.BaseDirectory, $"output");

            // Creates output folder if it doesn't exists
            Directory.CreateDirectory(outputFolder);

            string excelFullName = Path.Combine(outputFolder, fileName);

            if (GenerateExcel(excelFullName, reportModels))
            {
                return excelFullName;
            }
            else
            {
                return "";
            }
        }

        private int GetBreakfastInclusionValue(HotelRate hotelRate)
        {
            var rateTag = hotelRate.RateTags.FirstOrDefault(x => x.Name.ToLower() == "breakfast");

            if (rateTag != null)
            {
                return rateTag.Shape ? 1 : 0;
            }
            else
            {
                return 0;
            }
        }

        int colCount = 0;
        int rowCount = 0;

        private bool GenerateExcel(string fileName, List<ReportModel> reportData)
        {
            try
            {
                Console.WriteLine("Starting excel generation");

                using (var spreadsheetDoc = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookPart = spreadsheetDoc.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet();

                    FreezeTopRow(worksheetPart);

                    FillDataInWorksheet(reportData, worksheetPart);
                    worksheetPart.Worksheet.Save();

                    Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                    Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };

                    sheets.AppendChild(sheet);

                    AutoFilter autoFilter = new AutoFilter();
                    autoFilter.Reference = $"{GetExcelColumnNameByIndex(1)}1:{GetExcelColumnNameByIndex(colCount)}1"; //"A1:G1";
                    worksheetPart.Worksheet.Append(autoFilter);
                    SetConditionalFormatingStyle(workbookPart, worksheetPart, spreadsheetDoc);

                    workbookPart.Workbook.Save();

                    Console.WriteLine("Excel created successfully.");

                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        private static void FreezeTopRow(WorksheetPart worksheetPart)
        {
            SheetViews sheetViews1 = new SheetViews();

            SheetView sheetView1 = new SheetView()
            {
                TabSelected = true,
                WorkbookViewId = (UInt32Value)0U
            };
            Pane pane1 = new Pane()
            {
                VerticalSplit = 1D,
                TopLeftCell = "A2",
                ActivePane = PaneValues.BottomLeft,
                State = PaneStateValues.Frozen
            };

            sheetView1.Append(pane1);
            sheetViews1.Append(sheetView1);
            worksheetPart.Worksheet.Append(sheetViews1);
        }

        private void FillDataInWorksheet(List<ReportModel> reportData, WorksheetPart worksheetPart)
        {
            var sheetData = new SheetData();

            // Convert the list to datatable for easy access to column name and traversing by index/colName
            var reportDataTable = reportData.ToDataTable();

            Columns columns = new Columns();

            Row headerRow = new Row();

            Dictionary<String, CellValues> columnType = new Dictionary<string, CellValues>();
            UInt32Value i = 1;
            foreach (var props in typeof(ReportModel).GetProperties())
            {
                CellValues cellDataType =
                    props.PropertyType.Name.ToLower() == "decimal" || props.PropertyType.Name.ToLower().StartsWith("int") ? CellValues.Number :
                    props.PropertyType.Name.ToLower() == "datetime" ? CellValues.Date :
                    CellValues.String;
                columnType.Add(props.Name,
                    cellDataType
                    );

                Cell cell = new Cell
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue(props.Name.ToUpper()),
                };

                headerRow.AppendChild(cell);
                columns.Append(new Column
                {
                    Min = i,
                    Max = i++,
                    Width = props.Name == "RateName" ? 40D : 25D,
                    BestFit = true,
                    CustomWidth = true
                });
            }

            worksheetPart.Worksheet.Append(columns);
            sheetData.AppendChild(headerRow);

            foreach (DataRow hotelRateItem in reportDataTable.Rows)
            {
                Row newRow = new Row();
                foreach (var col in columnType)
                {
                    Cell cell = new Cell
                    {
                        //DataType = col.Value,
                        //CellValue = new CellValue(hotelRateItem[col.Key].ToString())
                    };

                    switch (col.Value)
                    {
                        case CellValues.Date:
                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(hotelRateItem[col.Key].ToString());
                            //cell.CellValue = new CellValue(DateTime.Parse(hotelRateItem[col.Key].ToString()));
                            //cell.StyleIndex = 1U;
                            break;
                        case CellValues.Number:
                        case CellValues.String:
                            cell.DataType = col.Value;
                            cell.CellValue = new CellValue(hotelRateItem[col.Key].ToString());
                            break;
                        default:
                            break;
                    }

                    newRow.AppendChild(cell);
                }

                sheetData.AppendChild(newRow);
            }
            worksheetPart.Worksheet.Append(sheetData);
        }

        private void SetConditionalFormatingStyle(WorkbookPart workbookPart, WorksheetPart worksheetPart, SpreadsheetDocument document)
        {
            // Using conditional formating to set the alternating row background

            WorkbookStylesPart stylesPart = document.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().FirstOrDefault();
            if (stylesPart == null)
            {
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet();
            }

            ConditionalFormatting conditionalFormatting1 = new ConditionalFormatting()
            {
                SequenceOfReferences = new ListValue<StringValue>()
                {
                    InnerText = $"{GetExcelColumnNameByIndex(1)}1:{GetExcelColumnNameByIndex(colCount)}{rowCount + 1}"//"A1:G105"
                }
            };

            ConditionalFormattingRule conditionalFormattingRule1 = new ConditionalFormattingRule()
            {
                Type = ConditionalFormatValues.Expression,
                FormatId = (UInt32Value)0U,
                Priority = 1
            };
            Formula formula1 = new Formula();
            formula1.Text = "MOD(ROW(),2)=0";

            conditionalFormattingRule1.Append(formula1);

            conditionalFormatting1.Append(conditionalFormattingRule1);

            worksheetPart.Worksheet.Append(conditionalFormatting1);

            /// Was trying to add date format.
            //SetNumberingFormats(stylesPart);

            stylesPart.Stylesheet.Append(GetFonts());
            stylesPart.Stylesheet.Append(GetDifferentialFormats());
        }

        public DifferentialFormats GetDifferentialFormats()
        {
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)3U };

            DifferentialFormat differentialFormat1 = new DifferentialFormat();

            Fill fill1 = new Fill();

            PatternFill patternFill1 = new PatternFill();
            BackgroundColor backgroundColor1 = new BackgroundColor() { Theme = (UInt32Value)4U, Tint = 0.59996337778862885D };

            patternFill1.Append(backgroundColor1);

            fill1.Append(patternFill1);

            differentialFormat1.Append(fill1);

            DifferentialFormat differentialFormat2 = new DifferentialFormat();

            Fill fill2 = new Fill();

            PatternFill patternFill2 = new PatternFill();
            BackgroundColor backgroundColor2 = new BackgroundColor() { Theme = (UInt32Value)4U, Tint = 0.59996337778862885D };

            patternFill2.Append(backgroundColor2);

            fill2.Append(patternFill2);

            differentialFormat2.Append(fill2);

            DifferentialFormat differentialFormat3 = new DifferentialFormat();

            //Font font1 = new Font();
            //Color color1 = new Color() { Theme = (UInt32Value)4U, Tint = -0.499984740745262D };

            //font1.Append(color1);

            Fill fill3 = new Fill();

            PatternFill patternFill3 = new PatternFill();
            BackgroundColor backgroundColor3 = new BackgroundColor() { Theme = (UInt32Value)4U, Tint = 0.59996337778862885D };

            patternFill3.Append(backgroundColor3);

            fill3.Append(patternFill3);

            //differentialFormat3.Append(font1);
            differentialFormat3.Append(fill3);

            differentialFormats1.Append(differentialFormat1);
            differentialFormats1.Append(differentialFormat2);
            differentialFormats1.Append(differentialFormat3);
            return differentialFormats1;
        }

        public void SetNumberingFormats(WorkbookStylesPart stylesPart)
        {
            NumberingFormats numberingFormats1 = new NumberingFormats() { Count = (UInt32Value)1U };
            NumberingFormat numberingFormat1 = new NumberingFormat() { NumberFormatId = (UInt32Value)165U, FormatCode = "[$-14009]d\\.m\\.yy;@" };

            numberingFormats1.Append(numberingFormat1);


            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)2U };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)165U, ApplyNumberFormat = true };

            cellFormats1.Append(cellFormat3);

            stylesPart.Stylesheet.Append(numberingFormats1);
            stylesPart.Stylesheet.Append(cellStyleFormats1);
            stylesPart.Stylesheet.Append(cellFormats1);

        }

        public Fonts GetFonts()
        {
            Fonts fonts = new Fonts() { Count = (UInt32Value)2U, KnownFonts = true };

            Font font = new Font();
            FontSize fontSize2 = new FontSize() { Val = 11D };
            Color color2 = new Color() { Theme = (UInt32Value)4U, Rgb = "#002060" };//, Tint = -0.249977111117893D };
            FontName fontName2 = new FontName() { Val = "Calibri" };
            //FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            //FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font.Append(fontSize2);
            font.Append(color2);
            font.Append(fontName2);
            //font.Append(fontFamilyNumbering2);
            //font.Append(fontScheme2);

            //fonts1.Append(font1);
            fonts.Append(font);
            return fonts;

        }

        private string GetExcelColumnNameByIndex(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = "";
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
    }
}
