using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

internal class Validation
{
    internal const string RU_X = "Х";
    internal const string EN_X = "X";
    internal const string EIGHT = "8";

    private readonly string _filePath1;
    private readonly string _filePath2;

    private readonly int _month;
    private readonly int _year;

    internal readonly Table Table1;
    private readonly Table _table2;

    internal readonly IEnumerable<string> NamesWorkedLastDayMonth;

    public Validation()
    {
        _filePath1 = GetFilePath("1.docx");
        _filePath2 = GetFilePath("2.docx");

        Table1 = GetLastTable(_filePath1);
        _table2 = GetLastTable(_filePath2);

        string monthName = GetStringsFromParagraph(_filePath1)[^3];

        _month = Array.IndexOf(DateTimeFormatInfo.CurrentInfo.MonthNames, monthName) + 1;

        _year = int.Parse(GetStringsFromParagraph(_filePath1)[^2]);

        NamesWorkedLastDayMonth = GetNamesWorkedLastDayMonth();
    }

    private static string[] GetStringsFromParagraph(string filePath1) =>
        GetFirstParagraph(filePath1).InnerText.Split(' ', StringSplitOptions.RemoveEmptyEntries);

    internal void CheckDaysInMonth()
    {
        string? lastDayMonth = Table1.Elements<TableRow>().First().Elements<TableCell>().Last().
            InnerText;

        int daysInMonth = DateTime.DaysInMonth(_year, _month);

        if (int.Parse(lastDayMonth) != daysInMonth)
            Console.WriteLine("days in month");
    }

    internal void CheckWeekendsColor()
    {
        bool HasIncorrectWeekendsColor = Table1.Elements<TableRow>().First().Elements<TableCell>().
            Any(cells => cells.Elements<TableCellProperties>().ElementAt(0).Shading is not null &&
            new DateOnly(_year, _month, int.Parse(cells.InnerText)).
            DayOfWeek is not (DayOfWeek.Saturday or DayOfWeek.Sunday));

        if (HasIncorrectWeekendsColor)
            Console.WriteLine("weekends color");
    }

    internal void CheckXsCount()
    {
        IEnumerable<List<TableCell>> rows = Table1.Elements<TableRow>().Select(cell =>
            cell.Elements<TableCell>().Skip(3).ToList());

        for (int i = 0; i < rows.First().Count; i++)
        {
            if (rows.Count(days => days[i].InnerText.EqualsOneOf(RU_X, EN_X)) is not (2 or 3))
                Console.WriteLine($"day {i + 1}");
        }
    }

    internal void CheckWeekendsWithEights()
    {
        bool hasWeekendsWithEights = Table1.Elements<TableRow>().ElementAt(3).
            Elements<TableCell>().Any(cells => cells.Elements<TableCellProperties>().ElementAt(0).
            Shading is not null && cells.InnerText.EqualsOneOf(EIGHT));

        if (hasWeekendsWithEights)
            Console.WriteLine("weekend with eights");
    }

    internal void CheckFirstDay()
    {
        bool hasIncorrectFirstDay = NamesWorkedLastDayMonth.
            Intersect(Table1.Elements<TableRow>().
            Where(rows => rows.Elements<TableCell>().ElementAt(3).InnerText.
            EqualsOneOf(RU_X, EN_X, EIGHT)).
            Select(rows => rows.Elements<TableCell>().ElementAt(1).InnerText.Replace(" ", ""))).
            Any();

        if (hasIncorrectFirstDay)
            Console.WriteLine("first day");
    }

    internal IEnumerable<string> GetNamesWorkedLastDayMonth() => _table2.Elements<TableRow>().
            Where(rows => rows.Elements<TableCell>().Last().InnerText.EqualsOneOf(RU_X, EN_X)).
            Select(rows => Regex.Replace(rows.Elements<TableCell>().ElementAt(1).InnerText, @"\s+",
                string.Empty));

    internal void CheckOrderXand8()
    {
        var days = Table1.Elements<TableRow>().ElementAt(3).Elements<TableCell>().
            Select(cell => cell.InnerText).ToArray();

        for (int i = 3; i < days.Length - 1; i++)
        {
            if (days[i].EqualsOneOf(RU_X, EN_X) && days[i + 1].EqualsOneOf(EIGHT))
                Console.WriteLine("X and eight");
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

    internal static string GetFilePath(string fileName)
    {
        string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
        string file = Path.Combine(currentDirectory, @"../../../static/" + fileName);
        string filePath = Path.GetFullPath(file);

        return filePath;
    }
}

public static class StringExtension
{
    public static bool EqualsOneOf(this string s, params string[] strings) =>
        strings.Contains(s.Trim(), StringComparer.OrdinalIgnoreCase);
}