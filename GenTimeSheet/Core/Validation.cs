using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GenTimeSheet.Core;

internal class Validation
{
    private const string DAYS_IN_MONTH_ERROR = "Неверное количество дней месяца!";
    private const string WEEKENDS_COLOR_ERROR = "Неверный цвет выходных!";
    private const string X_COUNT_ERROR = "Неверное количество людей в дне ";
    private const string WEEKENDS_WITH_EIGHTS_ERROR = "Восьмёрки на выходных!";
    private const string FIRST_DAY_ERROR = "Неверный Х в первый день!";
    private const string EIGHTS_NOT_EXIST_ERROR = "Нет восьмёрок";
    private const string EIGHTS_AFTER_X_ERROR = "Неверная восьмёрка после Х в день ";

    internal int Month { get; }
    internal int Year { get; }

    internal Table CurrentMonthTable { get; }

    private readonly Table _previousMonthTable;

    internal IEnumerable<string> NamesWorkedLastDayMonth { get; }

    internal List<string> ValidationErrors = [];

    internal Validation(string filePath1, string filePath2)
    {
        var parseDoc1 = new ParseDoc(filePath1);
        var parseDoc2 = new ParseDoc(filePath2);

        CurrentMonthTable = parseDoc1.GetTableFromEnd(1);
        _previousMonthTable = parseDoc2.GetTableFromEnd(1);

        string monthName = parseDoc1.GetStringsFromParagraph()[^3].ToLower();

        Month = Array.IndexOf(DateTimeFormatInfo.CurrentInfo.MonthNames, monthName) + 1;
        Year = int.Parse(parseDoc1.GetStringsFromParagraph()[^2]);

        NamesWorkedLastDayMonth = parseDoc2.GetNamesWorkedLastDayMonth(_previousMonthTable);
    }

    internal async Task ValidateDocx()
    {
        CheckDaysInMonth();
        await CheckWeekendsColor();
        CheckXsCount();
        CheckWeekendsWithEights();
        CheckFirstDay();
        CheckOrderXsAndEights();
    }    

    private void CheckDaysInMonth()
    {
        string? lastDayMonth = CurrentMonthTable.Elements<TableRow>().First().Elements<TableCell>().Last().
            InnerText;

        int daysInMonth = DateTime.DaysInMonth(Year, Month);

        if (int.Parse(lastDayMonth) != daysInMonth)
            ValidationErrors.Add(DAYS_IN_MONTH_ERROR);
    }

    private async Task CheckWeekendsColor()
    {
        List<int> weekends = await new CalendarHandler().GetMonthWeekends(Month - 1);

        var weekendsColor = CurrentMonthTable.Elements<TableRow>().First().Elements<TableCell>().
            Where(cells => cells.Elements<TableCellProperties>().ElementAt(0).Shading is not null).
            Select(s => int.Parse(s.InnerText));

        bool firstNotSecond = weekendsColor.Except(weekends).Any();
        bool secondNotFirst = weekends.Except(weekendsColor).Any();

        if (firstNotSecond || secondNotFirst)
            ValidationErrors.Add(WEEKENDS_COLOR_ERROR);
    }

    private void CheckXsCount()
    {
        IEnumerable<List<TableCell>> rows = CurrentMonthTable.Elements<TableRow>().Select(cell =>
            cell.Elements<TableCell>().Skip(3).ToList());

        for (int i = 0; i < rows.First().Count; i++)
        {
            if (rows.Count(days => days[i].InnerText.EqualsOneOf(Constants.RU_X, Constants.EN_X)) is not (2 or 3))
                ValidationErrors.Add($"{X_COUNT_ERROR}{i + 1}");
        }
    }

    private void CheckWeekendsWithEights()
    {
        bool hasWeekendsWithEights = CurrentMonthTable.Elements<TableRow>().ElementAt(3).
            Elements<TableCell>().Any(cells => cells.Elements<TableCellProperties>().ElementAt(0).
            Shading is not null && cells.InnerText.EqualsOneOf(Constants.EIGHT));

        if (hasWeekendsWithEights)
            ValidationErrors.Add(WEEKENDS_WITH_EIGHTS_ERROR);
    }

    private void CheckFirstDay()
    {
        bool hasIncorrectFirstDay = NamesWorkedLastDayMonth.
            Intersect(CurrentMonthTable.Elements<TableRow>().
            Where(rows => rows.Elements<TableCell>().ElementAt(3).InnerText.
            EqualsOneOf(Constants.RU_X, Constants.EN_X, Constants.EIGHT)).
            Select(rows => rows.Elements<TableCell>().ElementAt(1).InnerText.Replace(" ", ""))).
            Any();

        if (hasIncorrectFirstDay)
            ValidationErrors.Add(FIRST_DAY_ERROR);
    }

    private IEnumerable<string> GetNamesWorkedLastDayMonth() => _previousMonthTable.Elements<TableRow>().
            Where(rows => rows.Elements<TableCell>().Last().InnerText.EqualsOneOf(Constants.RU_X, Constants.EN_X)).
            Select(rows => Regex.Replace(rows.Elements<TableCell>().ElementAt(1).InnerText, @"\s+",
                string.Empty)).ToArray();

    private void CheckOrderXsAndEights()
    {
        var days = CurrentMonthTable.Elements<TableRow>().ElementAt(3).Elements<TableCell>().
            Select(cell => cell.InnerText).ToArray();

        if (days.All(cell => cell != Constants.EIGHT))
        {
            ValidationErrors.Add(EIGHTS_NOT_EXIST_ERROR);

            return;
        }

        for (int i = 3; i < days.Length - 1; i++)
        {
            if (days[i].EqualsOneOf(Constants.RU_X, Constants.EN_X) && days[i + 1].EqualsOneOf(Constants.EIGHT))
                ValidationErrors.Add(EIGHTS_AFTER_X_ERROR + $"{i - 1}");
        }
    }
}

internal static class StringExtension
{
    internal static bool EqualsOneOf(this string s, params string[] strings) =>
        strings.Contains(s.Trim(), StringComparer.OrdinalIgnoreCase);
}