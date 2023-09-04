using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

internal class WordValidation
{
    private string _filePath1;
    private string _filePath2;

    private int _month;
    private int _year;

    private Table _table1;
    private Table _table2;

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

    internal void CheckPeopleCount()
    {
        var schedule = new List<List<string>>();

        foreach (var row in _table1.Elements<TableRow>())
        {
            var name = new List<string>() { row.ElementAt(2).InnerText };
            var days = row.Elements<TableCell>().Skip(3).Select(c => c.InnerText);

            name.AddRange(days);
            schedule.Add(name);
        }

        for (int i = 1; i < schedule.Count; i++)
        {
            if (schedule.Count(r => r[i] is "Х" or "X") is not (2 or 3))
                Console.WriteLine($"day {i}");
        }
    }

    internal void CheckWeekendsWithEights()
    {
        bool hasWeekendsWithEights = _table1.Elements<TableRow>().ElementAt(3).Elements<TableCell>().
                                        Skip(3).
                                        Where(cells => cells.Elements<TableCellProperties>().ElementAt(0).Shading is not null && cells.InnerText == "8").
                                        Any();

        if (hasWeekendsWithEights)
            Console.WriteLine("weekend with eights");
    }

    internal void CheckFirstDay()
    {
        var hasIncorrectFirstDay = _table2.Elements<TableRow>().Where(rows => rows.Elements<TableCell>().Last().InnerText is "Х" or "X").
                                        Select(rows => rows.Elements<TableCell>().ElementAt(1).InnerText).
                                        Intersect(_table1.Elements<TableRow>().
                                        Where(rows => rows.Elements<TableCell>().ElementAt(3).InnerText is "Х" or "X" or "8").
                                        Select(rows => rows.Elements<TableCell>().ElementAt(1).InnerText)).
                                        Any();

        if (hasIncorrectFirstDay)
            Console.WriteLine("first day");
    }

    internal void CheckOrderXand8()
    {
        var half = _table1.Elements<TableRow>().ElementAt(3).Elements<TableCell>().Skip(3).Select(cell => cell.InnerText).ToArray();

        for (int i = 0; i < half.Length; i++)
        {
            if (half[i] is "Х" or "X" && half[i + 1] is "8")
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