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

internal class Validation
{
    private readonly string _filePath1;
    private readonly string _filePath2;

    private readonly int _month;
    private readonly int _year;

    internal readonly Table Table1;
    private readonly Table _table2;

    internal readonly IEnumerable<string> NamesWorkedLastDayMonth;

    internal List<string> ValidationErrors = [];
    internal List<int> Holidays { get; private set; }

    internal Validation()
    {
        _filePath1 = FileHandler.GetFilePath("1.docx");
        _filePath2 = FileHandler.GetFilePath("2.docx");

        Table1 = GetLastTable(_filePath1);
        _table2 = GetLastTable(_filePath2);

        string monthName = GetStringsFromParagraph(_filePath1)[^3].ToLower();

        _month = Array.IndexOf(DateTimeFormatInfo.CurrentInfo.MonthNames, monthName) + 1;       

        _year = int.Parse(GetStringsFromParagraph(_filePath1)[^2]);

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

        int daysInMonth = DateTime.DaysInMonth(_year, _month);

        if (int.Parse(lastDayMonth) != daysInMonth)
            ValidationErrors.Add("days in month");
    }

    private async Task CheckWeekendsColor()
    {
        Holidays = await new CalendarHandler().GetMonthHolidaysDates(_month - 1);

        bool HasIncorrectWeekendsColor = Table1.Elements<TableRow>().First().Elements<TableCell>().
            Any(cells => cells.Elements<TableCellProperties>().ElementAt(0).Shading is not null &&
            (new DateOnly(_year, _month, int.Parse(cells.InnerText)).
            DayOfWeek is not (DayOfWeek.Saturday or DayOfWeek.Sunday) ||
            !Holidays.Contains(int.Parse(cells.InnerText))));

        if (HasIncorrectWeekendsColor)
            ValidationErrors.Add("weekends color");
    }

    // TODO
    private List<int> GetWeekends()
    {
        string monthName = DateTimeFormatInfo.CurrentInfo.MonthNames[_month - 1];

    //    Web.GetWeekends();

        return [];
    }

    private void CheckXsCount()
    {
        IEnumerable<List<TableCell>> rows = Table1.Elements<TableRow>().Select(cell =>
            cell.Elements<TableCell>().Skip(3).ToList());

        for (int i = 0; i < rows.First().Count; i++)
        {
            if (rows.Count(days => days[i].InnerText.EqualsOneOf(Constants.RU_X, Constants.EN_X)) is not (2 or 3))
                ValidationErrors.Add($"day {i + 1}");
        }
    }

    private void CheckWeekendsWithEights()
    {
        bool hasWeekendsWithEights = Table1.Elements<TableRow>().ElementAt(3).
            Elements<TableCell>().Any(cells => cells.Elements<TableCellProperties>().ElementAt(0).
            Shading is not null && cells.InnerText.EqualsOneOf(Constants.EIGHT));

        if (hasWeekendsWithEights)
            ValidationErrors.Add("weekend with eights");
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
            ValidationErrors.Add("first day");
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
            ValidationErrors.Add("eights is empty");

        for (int i = 3; i < days.Length - 1; i++)
        {
            if (days[i].EqualsOneOf(Constants.RU_X, Constants.EN_X) && days[i + 1].EqualsOneOf(Constants.EIGHT))
                ValidationErrors.Add("X and eight");
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