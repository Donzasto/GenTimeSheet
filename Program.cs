using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
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

void UpdateCell(string docName)
{
    using SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true);

    string id = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().GetFirstChild<Sheet>().
        Id.Value;

    WorksheetPart worksheetPart = (WorksheetPart)spreadSheet.WorkbookPart.GetPartById(id);

    Worksheet worksheet = worksheetPart.Worksheet;

    SheetData sheetData = worksheet.Elements<SheetData>().First();

    var namesColumn = worksheet.Descendants<Row>().Select(row => row.Elements<Cell>().ElementAt(1));

    DocumentFormat.OpenXml.Wordprocessing.Table table = wordValidation.Table1;

    var sst = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First().
        SharedStringTable;

    foreach (var cell in namesColumn)
    {
        if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
        {
            int ssid = int.Parse(cell.InnerText);

            string text = sst.ChildElements[ssid].InnerText;

            int rowIndex = int.Parse(Regex.Match(cell.CellReference, @"\d+").Value) - 1;

            IEnumerable<TableRow> personalDays = table.Elements<TableRow>().Where(row =>
                string.Compare(row.Elements<TableCell>().ElementAt(1).InnerText, text,
                CultureInfo.CurrentCulture, CompareOptions.IgnoreSymbols) == 0);

            if (!personalDays.Any())
            {
                continue;
            }

            var days = personalDays.First();

            int count = days.Count();

            for (int i = 0, cellIndex = 0; i < count; i++, cellIndex++)
            {
                const int passColumn = 19;

                if (cellIndex == passColumn)
                    cellIndex++;

                if (days.ElementAt(i).InnerText.Trim() is WordValidation.RU_X or WordValidation.EN_X)
                {
                    FIllCell(cellIndex, sheetData, rowIndex, "Ф", CellValues.String);
                    FIllCell(cellIndex, sheetData, rowIndex + 1, "16", CellValues.Number);
                    FIllCell(cellIndex, sheetData, rowIndex + 2, "2", CellValues.Number);

                    if (cellIndex + 1 == passColumn)
                        cellIndex++;

                    if (cellIndex == count)
                        break;

                    FIllCell(cellIndex + 1, sheetData, rowIndex, "Ф", CellValues.String);
                    FIllCell(cellIndex + 1, sheetData, rowIndex + 1, "8", CellValues.Number);
                    FIllCell(cellIndex + 1, sheetData, rowIndex + 2, "6", CellValues.Number);

                    if (cellIndex + 1 == 19)
                        cellIndex--;
                }
                else if (days.ElementAt(i).InnerText.Trim() is "О")
                {
                    FIllCell(cellIndex, sheetData, rowIndex, "О", CellValues.String);
                }
                else if (days.ElementAt(i).InnerText.Trim() is "Б")
                {
                    FIllCell(cellIndex, sheetData, rowIndex, "Б", CellValues.String);
                }
            }
        }
    }
}

void FIllCell(int cellIndex, SheetData sheetData, int rowIndex, string text, EnumValue<CellValues> dataType)
{
    IEnumerable<Cell> cells = sheetData.Elements<Row>().ElementAt(rowIndex).Elements<Cell>();
    cells.ElementAt(cellIndex).DataType = dataType;
    cells.ElementAt(cellIndex).CellValue = new CellValue() { Text = text };
}