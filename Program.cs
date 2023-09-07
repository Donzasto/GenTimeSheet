using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

WordValidation wordValidation = new();

wordValidation.CheckDaysInMonth();
wordValidation.CheckWeekendsColor();
wordValidation.CheckXsCount();
wordValidation.CheckWeekendsWithEights();
wordValidation.CheckFirstDay();
wordValidation.CheckOrderXand8();

string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
string file = Path.Combine(currentDirectory, @"../../../static/" + "1.xlsx");
string docName = Path.GetFullPath(file);

UpdateCell(docName);

static WorksheetPart GetWorksheetPartByName(SpreadsheetDocument document, string sheetName)
{
    IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().
                    Elements<Sheet>().Where(s => s.Name == sheetName);

    if (sheets.Count() == 0)
        return null;

    string relationshipId = sheets.First().Id.Value;

    WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);

    return worksheetPart;
}

void UpdateCell(string docName)
{
    using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
    {
        WorksheetPart worksheetPart = GetWorksheetPartByName(spreadSheet,
            spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().First().Name);

        if (worksheetPart != null)
        {
            var worksheet = worksheetPart.Worksheet;

            SheetData sheetData = worksheet.Elements<SheetData>().First();

            DocumentFormat.OpenXml.Wordprocessing.Table table = wordValidation.Table1;

            var sst = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First().SharedStringTable;

            var cells = worksheet.Descendants<Cell>();

            for (int i = 11; i < 62; i++)
            {
                if (i == 44)
                    continue;

                Cell cell = sheetData.Elements<Row>().ElementAt(i).Elements<Cell>().ElementAt(1);

                if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
                {
                    int ssid = int.Parse(cell.CellValue.Text);

                    string str = sst.ChildElements[ssid].InnerText;

                    var name = table.Elements<TableRow>().Where(row => string.Compare(row.Elements<TableCell>().ElementAt(1).InnerText, str, CultureInfo.CurrentCulture, CompareOptions.IgnoreSymbols) == 0);

                    if (!name.Any())
                    {
                        continue;
                    }

                    var n = name.First();

                    for (int k = 0, cellIndex = 0; k < n.Count(); k++, cellIndex++)
                    {
                        if (cellIndex == 19)
                            cellIndex++;
                        int rowIndex = i;
                        if (n.ElementAt(k).InnerText.Trim() is "Х")
                        {
                            Console.WriteLine(str + " " + $"{i}");
                            var c1 = sheetData.Elements<Row>().ElementAt(rowIndex).Elements<Cell>();
                            c1.ElementAt(cellIndex).
                                DataType = CellValues.String;
                            c1.ElementAt(cellIndex).
                                CellValue = new CellValue("Ф");

                            var c2 = sheetData.Elements<Row>().ElementAt(1 + rowIndex).Elements<Cell>();
                            c2.ElementAt(cellIndex).
                                DataType = CellValues.Number;
                            c2.ElementAt(cellIndex).
                                CellValue = new CellValue("16");


                            var c3 = sheetData.Elements<Row>().ElementAt(2 + rowIndex).Elements<Cell>();
                            c3.ElementAt(cellIndex).
                                DataType = CellValues.Number;
                            c3.ElementAt(cellIndex).
                                CellValue = new CellValue("[$Н/]#") { Text = "2" };

                            if (cellIndex + 1 == 19)
                                cellIndex++;

                            if (cellIndex == n.Count())
                                break;

                            var c1r = sheetData.Elements<Row>().ElementAt(rowIndex).Elements<Cell>();
                            c1r.ElementAt(cellIndex + 1).
                                DataType = CellValues.String;
                            c1r.ElementAt(cellIndex + 1).
                                CellValue = new CellValue("Ф");

                            var c2r = sheetData.Elements<Row>().ElementAt(1 + rowIndex).Elements<Cell>();
                            c2r.ElementAt(cellIndex + 1).
                                DataType = CellValues.Number;
                            c2r.ElementAt(cellIndex + 1).
                                CellValue = new CellValue("8");

                            var c3r = sheetData.Elements<Row>().ElementAt(2 + rowIndex).Elements<Cell>();
                            c3r.ElementAt(cellIndex + 1).
                                DataType = CellValues.Number;
                            c3r.ElementAt(cellIndex + 1).
                                CellValue = new CellValue("[$Н/]#") { Text = "6" };

                            if (cellIndex + 1 == 19)
                                cellIndex--;
                        }
                        else if (n.ElementAt(k).InnerText.Trim() is "О")
                        {
                            var c1 = sheetData.Elements<Row>().ElementAt(rowIndex).Elements<Cell>();
                            c1.ElementAt(cellIndex).
                                DataType = CellValues.String;
                            c1.ElementAt(cellIndex).
                                CellValue = new CellValue("О");
                        }
                        else if (n.ElementAt(k).InnerText.Trim() is "Б")
                        {
                            var c1 = sheetData.Elements<Row>().ElementAt(rowIndex).Elements<Cell>();
                            c1.ElementAt(cellIndex).
                                DataType = CellValues.String;
                            c1.ElementAt(cellIndex).
                                CellValue = new CellValue("Б");
                        }
                    }

                }
            }

            worksheetPart.Worksheet.Save();

            spreadSheet.WorkbookPart.Workbook.Save();
        }
    }
}