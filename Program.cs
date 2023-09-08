using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

Validation validation = new();

validation.CheckDaysInMonth();
validation.CheckWeekendsColor();
validation.CheckXsCount();
validation.CheckWeekendsWithEights();
validation.CheckFirstDay();
validation.CheckOrderXand8();

string docName = Validation.GetFilePath("1.xlsx");

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

    IEnumerable<TableRow> table = validation.Table1.Elements<TableRow>().Skip(1);

    var sst = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First().
        SharedStringTable;

    var u = sst.ChildElements;

    foreach (var cell in namesColumn)
    {
        if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
        {
            int ssid = int.Parse(cell.InnerText);

            string name = sst.ChildElements[ssid].InnerText;

            int rowIndex = int.Parse(Regex.Match(cell.CellReference, @"\d+").Value) - 1;

            IEnumerable<TableRow> personalDays = table.Where(row =>
                string.Compare(row.Elements<TableCell>().ElementAt(1).InnerText, name,
                CultureInfo.CurrentCulture, CompareOptions.IgnoreSymbols) == 0);

            if (!personalDays.Any())
            {
                continue;
            }

            var days = personalDays.First();

            int countDays = days.Count();


            if (validation.NamesWorkedLastDayMonth.
                Contains(Regex.Replace(name, @"\s+", string.Empty)))
            {
                FillCells(4, sheetData, rowIndex);
            }

            for (int i = 0, cellIndex = 0; i < countDays; i++, cellIndex++)
            {
                const int passColumn = 19;

                if (cellIndex == passColumn)
                    cellIndex++;


                if (days.ElementAt(i).InnerText.Trim() is Validation.RU_X or Validation.EN_X)
                {
                    FillCell(cellIndex, sheetData, rowIndex, "Ф", CellValues.String);
                    FillCell(cellIndex, sheetData, rowIndex + 1, "16", CellValues.Number);
                    FillCell(cellIndex, sheetData, rowIndex + 2, "2", CellValues.Number);

                    if (cellIndex + 1 == passColumn)
                        cellIndex++;

                    if (cellIndex == countDays)
                        break;

                    FillCells(cellIndex + 1, sheetData, rowIndex);

                    if (cellIndex + 1 == passColumn)
                        cellIndex--;
                }
                else if (days.ElementAt(i).InnerText.Trim() is "О")
                {
                    FillCell(cellIndex, sheetData, rowIndex, "О", CellValues.String);
                }
                else if (days.ElementAt(i).InnerText.Trim() is "Б")
                {
                    FillCell(cellIndex, sheetData, rowIndex, "Б", CellValues.String);
                }
            }

            RecalculateFormuls(19, sheetData, rowIndex);
            RecalculateFormuls(35, sheetData, rowIndex);
        }
    }
}

void FillCells(int cellIndex, SheetData sheetData, int rowIndex)
{
    FillCell(4, sheetData, rowIndex, "Ф", CellValues.String);
    FillCell(4, sheetData, rowIndex + 1, "8", CellValues.Number);
    FillCell(4, sheetData, rowIndex + 2, "6", CellValues.Number);
}

void FillCell(int cellIndex, SheetData sheetData, int rowIndex, string text,
    EnumValue<CellValues> dataType)
{
    IEnumerable<Cell> cells = sheetData.Elements<Row>().ElementAt(rowIndex).Elements<Cell>();
    Cell cell = cells.ElementAt(cellIndex);
    cell.DataType = dataType;
    cell.CellValue = new CellValue() { Text = text };
}

void RecalculateFormuls(int cellIndex, SheetData sheetData, int rowIndex)
{
    Row row = sheetData.Elements<Row>().ElementAt(rowIndex + 1);
    row.Elements<Cell>().ElementAt(cellIndex).CellValue?.Remove();

    row = sheetData.Elements<Row>().ElementAt(rowIndex + 2);
    row.Elements<Cell>().ElementAt(cellIndex).CellValue?.Remove();
}