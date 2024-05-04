using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GenTimeSheet.Core;

internal class Generator
{
    private IEnumerable<IEnumerable<Cell>> _tableSheet;

    private readonly Validation _validation;
    private List<int> _holidays;

    internal Generator(Validation validation)
    {
        _validation = validation;
    }

    private string GetMonthTemplatePath()
    {
        int days = DateTime.DaysInMonth(_validation.Year, _validation.Month);

        return "templates/" + days + ".xlsx";
    }

    private void SetTimesheetNumber(OpenXmlElement element)
    {
        string value = "Табель №" + _validation.Month * 2;

        element.RemoveAllChildren();
        element.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Text(value));

    }

    private void SetMonthHeader(OpenXmlElement element)
    {
        DateTimeFormatInfo dateTimeFormatInfo = new CultureInfo("ru-RU").DateTimeFormat;

        string monthGenitiveNames = dateTimeFormatInfo.MonthGenitiveNames[_validation.Month - 1];

        int days = DateTime.DaysInMonth(_validation.Year, _validation.Month);

        string value = $"C 1 {monthGenitiveNames} по {days} {monthGenitiveNames}";

        element.RemoveAllChildren();
        element.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Text(value));
    }

    internal async Task Run()
    {
        try
        {
            await GetHolidays();

            GenerateFile();
        }
        catch
        {
            throw;
        }
    }

    private void GenerateFile()
    {
        string templatePath = GetMonthTemplatePath();
        string filePath = FileHandler.GetFilePath(templatePath);
        string format = DateTime.Now.ToString("d-MM-yyyy-HH-mm-ss");

        var template = SpreadsheetDocument.Open(filePath, true);

        using SpreadsheetDocument spreadSheet = template.Clone($"../../../../GenTimeSheet/files/{format}.xlsx", true);

        template.Dispose();

        string? id = spreadSheet.WorkbookPart?.Workbook?.GetFirstChild<Sheets>()?.GetFirstChild<Sheet>()?.
            Id?.Value;

        if (id is null)
        {
            throw new Exception("The worksheet does not have an ID.");
        }

        WorksheetPart worksheetPart = (WorksheetPart)spreadSheet.WorkbookPart!.GetPartById(id);

        Worksheet worksheet = worksheetPart.Worksheet;

        IEnumerable<Cell> namesColumn = worksheet.Descendants<Row>().Select(row =>
            row.Elements<Cell>().ElementAt(1));

        IEnumerable<TableRow> table = _validation.CurrentMonthTable.Elements<TableRow>().Skip(1);

        SharedStringTable sharedStringTable = spreadSheet.WorkbookPart.
            GetPartsOfType<SharedStringTablePart>().First().SharedStringTable;

        SetTimesheetNumber(sharedStringTable.ElementAt(0));
        SetMonthHeader(sharedStringTable.ElementAt(4));

        _tableSheet = worksheet.GetFirstChild<SheetData>().Elements<Row>().
            Select(row => row.Elements<Cell>().Skip(4).Take(15).
            Concat(row.Elements<Cell>().Skip(19))).ToArray();

        PopulateCells(namesColumn, table, sharedStringTable);
    }

    private void PopulateCells(IEnumerable<Cell> namesColumn, IEnumerable<TableRow> table, SharedStringTable sharedStringTable)
    {
        foreach (var cell in namesColumn)
        {
            int formulsColumn = 15;

            if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
            {
                int ssid = int.Parse(cell.InnerText);

                string name = sharedStringTable.ChildElements[ssid].InnerText;

                int rowIndex = int.Parse(Regex.Match(cell.CellReference, @"\d+").Value) - 1;

                IEnumerable<TableRow> personalDays = table.Where(row =>
                    string.Compare(row.Elements<TableCell>().ElementAt(1).InnerText, name,
                    CultureInfo.CurrentCulture, CompareOptions.IgnoreSymbols) == 0);

                if (!personalDays.Any())
                {
                    continue;
                }

                TableRow days = personalDays.First();

                var markedDays = days.Skip(4).Select((Value, Index) => new { Value, Index }).
                    Where(cell => cell.Value.InnerText.Any());

                int daysCount = days.Count() - 4;

                foreach (var day in markedDays)
                {
                    int cellIndex = day.Index;
                    string innerText = day.Value.InnerText;

                    if (cellIndex >= formulsColumn)
                        cellIndex++;

                    if (innerText == Constants.RU_B)
                    {
                        SetCell(cellIndex, rowIndex, Constants.RU_B, CellValues.String);
                    }
                    else if (innerText.EqualsOneOf(Constants.RU_X, Constants.EN_X))
                    {
                        string dayStatus = GetDayStatus(cellIndex);

                        SetCell(cellIndex, rowIndex, dayStatus, CellValues.String);
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
                }

                RecalculateFormuls(formulsColumn, rowIndex);
                formulsColumn = daysCount + 1;
                RecalculateFormuls(formulsColumn, rowIndex);
            }
        }
    }

    private async Task GetHolidays()
    {
        _holidays = await new CalendarHandler().GetMonthHolidaysDates(_validation.Month - 1);
    }

    private string GetDayStatus(int dayIndex)
    {
        string dayStatus = Constants.RU_F;

        if (dayIndex < 15)
        {
            ++dayIndex;
        }

        if (_holidays.Contains(dayIndex))
        {
            dayStatus = Constants.RU_RP;
        }

        return dayStatus;
    }

    private void SetCells(int cellIndex, int rowIndex)
    {
        string cellContent = GetDayStatus(cellIndex);

        SetCell(cellIndex, rowIndex, cellContent, CellValues.String);
        SetCell(cellIndex, rowIndex + 1, Constants.EIGHT, CellValues.Number);
        SetCell(cellIndex, rowIndex + 2, Constants.SIX, CellValues.Number);
    }

    private void SetCell(int cellIndex, int rowIndex, string text,
        EnumValue<CellValues> dataType)
    {
        IEnumerable<Cell> cells = _tableSheet.Select(row => row).ElementAt(rowIndex);

        Cell cell = cells.ElementAt(cellIndex);
        cell.DataType = dataType;
        cell.CellValue = new CellValue() { Text = text };
    }

    private void RecalculateFormuls(int cellIndex, int rowIndex)
    {
        IEnumerable<Cell> cells = _tableSheet.Select(row => row).ElementAt(rowIndex + 1);
        cells.ElementAt(cellIndex).CellValue?.Remove();

        cells = _tableSheet.Select(row => row).ElementAt(rowIndex + 2);
        cells.ElementAt(cellIndex).CellValue?.Remove();
    }
}