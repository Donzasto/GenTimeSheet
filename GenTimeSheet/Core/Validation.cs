using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GenTimeSheet.Core;

public class Validation
{
    private readonly string _filePath1;
    private readonly string _filePath2;

    public int Month { get; private set; }
    public int Year { get; private set; }

    internal readonly Table Table1;
    private readonly Table _table2;

    internal readonly IEnumerable<string> NamesWorkedLastDayMonth;

    public List<string> ValidationErrors = [];

    private const string DAYS_IN_MONTH_ERROR = "Неверное количество дней месяца!";
    private const string WEEKENDS_COLOR_ERROR = "Неверный цвет выходных!";
    private const string X_COUNT_ERROR = "Неверное количество людей в дне ";
    private const string WEEKENDS_WITH_EIGHTS_ERROR = "Восьмёрки на выходных!";
    private const string FIRST_DAY_ERROR = "Неверный Х в первый день!";
    private const string EIGHTS_NOT_EXIST_ERROR = "Нет восьмёрок";
    private const string EIGHTS_AFTER_X_ERROR = "Неверная восьмёрка после Х в день ";

    public Validation(string filePath1, string filePath2)
    {
        _filePath1 = FileHandler.GetFilePath(filePath1);
        _filePath2 = FileHandler.GetFilePath(filePath2);

        Table1 = GetLastTable(_filePath1);
        _table2 = GetLastTable(_filePath2);
        
        string monthName = GetStringsFromParagraph(_filePath1)[^3].ToLower();

        Month = Array.IndexOf(DateTimeFormatInfo.CurrentInfo.MonthNames, monthName) + 1;
        
        Year = int.Parse(GetStringsFromParagraph(_filePath1)[^2]);        

        NamesWorkedLastDayMonth = GetNamesWorkedLastDayMonth();
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

    private static string[] GetStringsFromParagraph(string filePath1) =>
        GetFirstParagraph(filePath1).InnerText.Split(' ', StringSplitOptions.RemoveEmptyEntries);

    private void CheckDaysInMonth()
    {
        string? lastDayMonth = Table1.Elements<TableRow>().First().Elements<TableCell>().Last().
            InnerText;

        int daysInMonth = DateTime.DaysInMonth(Year, Month);

        if (int.Parse(lastDayMonth) != daysInMonth)
            ValidationErrors.Add(DAYS_IN_MONTH_ERROR);
    }

    private async Task CheckWeekendsColor()
    {
        List<int> weekends = await new CalendarHandler().GetMonthWeekends(Month - 1);

        var weekendsColor = Table1.Elements<TableRow>().First().Elements<TableCell>().
            Where(cells => cells.Elements<TableCellProperties>().ElementAt(0).Shading is not null).
            Select(s => int.Parse(s.InnerText));

        bool firstNotSecond = weekendsColor.Except(weekends).Any();
        bool secondNotFirst = weekends.Except(weekendsColor).Any();

        if (firstNotSecond || secondNotFirst)
            ValidationErrors.Add(WEEKENDS_COLOR_ERROR);
    }

    private void CheckXsCount()
    {
        IEnumerable<List<TableCell>> rows = Table1.Elements<TableRow>().Select(cell =>
            cell.Elements<TableCell>().Skip(3).ToList());

        for (int i = 0; i < rows.First().Count; i++)
        {
            if (rows.Count(days => days[i].InnerText.EqualsOneOf(Constants.RU_X, Constants.EN_X)) is not (2 or 3))
                ValidationErrors.Add($"{X_COUNT_ERROR}{i + 1}");
        }
    }

    private void CheckWeekendsWithEights()
    {
        bool hasWeekendsWithEights = Table1.Elements<TableRow>().ElementAt(3).
            Elements<TableCell>().Any(cells => cells.Elements<TableCellProperties>().ElementAt(0).
            Shading is not null && cells.InnerText.EqualsOneOf(Constants.EIGHT));

        if (hasWeekendsWithEights)
            ValidationErrors.Add(WEEKENDS_WITH_EIGHTS_ERROR);
    }

    private void CheckFirstDay()
    {
        bool hasIncorrectFirstDay = NamesWorkedLastDayMonth.
            Intersect(Table1.Elements<TableRow>().
            Where(rows => rows.Elements<TableCell>().ElementAt(3).InnerText.
            EqualsOneOf(Constants.RU_X, Constants.EN_X, Constants.EIGHT)).
            Select(rows => rows.Elements<TableCell>().ElementAt(1).InnerText.Replace(" ", ""))).
            Any();

        if (hasIncorrectFirstDay)
            ValidationErrors.Add(FIRST_DAY_ERROR);
    }

    private IEnumerable<string> GetNamesWorkedLastDayMonth() => _table2.Elements<TableRow>().
            Where(rows => rows.Elements<TableCell>().Last().InnerText.EqualsOneOf(Constants.RU_X, Constants.EN_X)).
            Select(rows => Regex.Replace(rows.Elements<TableCell>().ElementAt(1).InnerText, @"\s+",
                string.Empty)).ToArray();

    private void CheckOrderXsAndEights()
    {
        var days = Table1.Elements<TableRow>().ElementAt(3).Elements<TableCell>().
            Select(cell => cell.InnerText).ToArray();

        if(days.All(cell => cell != Constants.EIGHT))
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

    private static Body GetBody(string filePath)
    {
        using WordprocessingDocument wordprocessingDocument =
            WordprocessingDocument.Open(filePath, false);

        return wordprocessingDocument.MainDocumentPart.Document.Body;
    }

    internal Table GetLastTable(string filePath) => GetElements<Table>(filePath).Last();

    private static Paragraph GetFirstParagraph(string filePath) =>
        GetElements<Paragraph>(filePath).First();

    private static IEnumerable<T> GetElements<T>(string filePath) where T : OpenXmlElement =>
        GetBody(filePath).Elements<T>();
}

public static class StringExtension
{
    public static bool EqualsOneOf(this string s, params string[] strings) =>
        strings.Contains(s.Trim(), StringComparer.OrdinalIgnoreCase);
}