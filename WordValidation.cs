using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

internal class WordValidation
{
    private const string RU_X = "Х";
    private const string EN_X = "X";
    private const string EIGHT = "8";

    private readonly string _filePath1;
    private readonly string _filePath2;

    private readonly int _month;
    private readonly int _year;

    private readonly Table _table1;
    private readonly Table _table2;

    public WordValidation()
    {
        _filePath1 = GetFilePath("1.docx");
        _filePath2 = GetFilePath("2.docx");

        _table1 = GetLastTable(_filePath1);
        _table2 = GetLastTable(_filePath2);

        string monthName = GetStringsFromParagraph(_filePath1)[^3];

        _month = Array.IndexOf(DateTimeFormatInfo.CurrentInfo.MonthNames, monthName) + 1;

        _year = int.Parse(GetStringsFromParagraph(_filePath1)[^2]);
    }

    string[] GetStringsFromParagraph(string filePath1) =>
                GetFirstParagraph(filePath1).InnerText.Split(' ', StringSplitOptions.RemoveEmptyEntries);

    internal void CheckDaysInMonth()
    {
        string lastDayMonth = _table1.Elements<TableRow>().ElementAt(0).Elements<TableCell>().Last().InnerText;

        int daysInMonth = DateTime.DaysInMonth(_year, _month);

        if (int.Parse(lastDayMonth) != daysInMonth)
            Console.WriteLine("days in month");
    }

    internal void CheckWeekendsColor()
    {
        bool HasIncorrectWeekendsColor = _table1.Elements<TableRow>().ElementAt(0).Elements<TableCell>().
                                Skip(3).
                                Where(cells => cells.Elements<TableCellProperties>().ElementAt(0).Shading is not null &&
                                    new DateOnly(_year, _month, int.Parse(cells.InnerText)).DayOfWeek is not (DayOfWeek.Saturday or DayOfWeek.Sunday)).
                                Any();

        if (HasIncorrectWeekendsColor)
            Console.WriteLine("weekends color");
    }

    internal void CheckXsCount()
    {
        var schedule = new List<List<string>>();

        foreach (var row in _table1.Elements<TableRow>())
        {
            var days = row.Elements<TableCell>().Skip(3).Select(c => c.InnerText).ToList();

            schedule.Add(days);
        }

        for (int i = 0; i < schedule[0].Count; i++)
        {
            if (schedule.Count(r => r[i] is RU_X or EN_X) is not (2 or 3))
                Console.WriteLine($"day {i + 1}");
        }
    }

    internal void CheckWeekendsWithEights()
    {
        bool hasWeekendsWithEights = _table1.Elements<TableRow>().ElementAt(3).Elements<TableCell>().
                                        Skip(3).
                                        Where(cells => cells.Elements<TableCellProperties>().ElementAt(0).Shading is not null && cells.InnerText is EIGHT).
                                        Any();

        if (hasWeekendsWithEights)
            Console.WriteLine("weekend with eights");
    }

    internal void CheckFirstDay()
    {
        var hasIncorrectFirstDay = _table2.Elements<TableRow>().Where(rows => rows.Elements<TableCell>().Last().InnerText is RU_X or EN_X).
                                        Select(rows => rows.Elements<TableCell>().ElementAt(1).InnerText).
                                        Intersect(_table1.Elements<TableRow>().
                                        Where(rows => rows.Elements<TableCell>().ElementAt(3).InnerText is RU_X or EN_X or EIGHT).
                                        Select(rows => rows.Elements<TableCell>().ElementAt(1).InnerText)).
                                        Any();

        if (hasIncorrectFirstDay)
            Console.WriteLine("first day");
    }

    internal void CheckOrderXand8()
    {
        var half = _table1.Elements<TableRow>().ElementAt(3).Elements<TableCell>().Skip(3).Select(cell => cell.InnerText).ToArray();

        for (int i = 0; i < half.Length - 1; i++)
        {
            if (half[i] is RU_X or EN_X && half[i + 1] is EIGHT)
                Console.WriteLine("X and eight");
        }
    }

    private Body GetBody(string filePath)
    {
        using WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filePath, false);

        return wordprocessingDocument.MainDocumentPart.Document.Body;
    }

    private Table GetLastTable(string filePath) => GetElements<Table>(filePath).Last();

    private Paragraph GetFirstParagraph(string filePath) => GetElements<Paragraph>(filePath).First();

    private IEnumerable<T> GetElements<T>(string filePath) where T : OpenXmlElement => GetBody(filePath).Elements<T>();

    private string GetFilePath(string fileName)
    {
        string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
        string file = Path.Combine(currentDirectory, @"../../../static/" + fileName);
        string filePath = Path.GetFullPath(file);

        return filePath;
    }
}