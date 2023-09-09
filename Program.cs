using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

Validation validation = new();

validation.ValidateDocx();

string filepath = Validation.GetFilePath("1.xlsx");

IEnumerable<IEnumerable<Cell>> tableSheet;

UpdateCells(filepath);

void UpdateCells(string filepath)
{
    using SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filepath, true);

    string id = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().GetFirstChild<Sheet>().
        Id.Value;

    WorksheetPart worksheetPart = (WorksheetPart)spreadSheet.WorkbookPart.GetPartById(id);

    Worksheet worksheet = worksheetPart.Worksheet;

    IEnumerable<Cell> namesColumn = worksheet.Descendants<Row>().Select(row =>
        row.Elements<Cell>().ElementAt(1));

    IEnumerable<TableRow> table = validation.Table1.Elements<TableRow>().Skip(1);

    SharedStringTable sst = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().
        First().SharedStringTable;

    tableSheet = worksheet.GetFirstChild<SheetData>().Elements<Row>().
        Select(row => row.Elements<Cell>().Skip(4).Take(15).
        Concat(row.Elements<Cell>().Skip(19))).ToArray();

    foreach (var cell in namesColumn)
    {
        int formulsColumn = 15;

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

            TableRow days = personalDays.First();

            if (validation.NamesWorkedLastDayMonth.
                Contains(Regex.Replace(name, @"\s+", string.Empty)))
            {
                SetCells(0, rowIndex);
            }

            var markedDays = days.Skip(4).Select((Value, Index) => new { Value, Index }).
                Where(cell => cell.Value.InnerText.Any());

            int daysCount = days.Count() - 4;

            foreach (var day in markedDays)
            {
                int cellIndex = day.Index;
                string innerText = day.Value.InnerText;

                if (cellIndex >= formulsColumn)
                    cellIndex++;

                if (innerText.EqualsOneOf(Constants.RU_X, Constants.EN_X))
                {
                    SetCell(cellIndex, rowIndex, Constants.RU_F, CellValues.String);
                    SetCell(cellIndex, rowIndex + 1, Constants.SIXTEEN, CellValues.Number);
                    SetCell(cellIndex, rowIndex + 2, Constants.TWO, CellValues.Number);

                    if (cellIndex == daysCount)
                        break;

                    if (cellIndex + 1 == formulsColumn)
                        cellIndex++;

                    SetCells(cellIndex + 1, rowIndex);
                }
                else if (innerText.EqualsOneOf(Constants.RU_O, Constants.EN_O))
                {
                    SetCell(cellIndex, rowIndex, Constants.RU_O, CellValues.String);
                }
                else if (innerText.EqualsOneOf(Constants.RU_B))
                {
                    SetCell(cellIndex, rowIndex, Constants.RU_B, CellValues.String);
                }
            }

            RecalculateFormuls(formulsColumn, rowIndex);
            formulsColumn = daysCount + 1;
            RecalculateFormuls(formulsColumn, rowIndex);
        }
    }
}

void SetCells(int cellIndex, int rowIndex)
{
    SetCell(cellIndex, rowIndex, Constants.RU_F, CellValues.String);
    SetCell(cellIndex, rowIndex + 1, Constants.EIGHT, CellValues.Number);
    SetCell(cellIndex, rowIndex + 2, Constants.SIX, CellValues.Number);
}

void SetCell(int cellIndex, int rowIndex, string text,
    EnumValue<CellValues> dataType)
{
    IEnumerable<Cell> cells = tableSheet.Select(row => row).ElementAt(rowIndex);

    Cell cell = cells.ElementAt(cellIndex);
    cell.DataType = dataType;
    cell.CellValue = new CellValue() { Text = text };
}

void RecalculateFormuls(int cellIndex, int rowIndex)
{
    IEnumerable<Cell> cells = tableSheet.Select(row => row).ElementAt(rowIndex + 1);
    cells.ElementAt(cellIndex).CellValue?.Remove();

    cells = tableSheet.Select(row => row).ElementAt(rowIndex + 2);
    cells.ElementAt(cellIndex).CellValue?.Remove();
}