using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

string filePath1 = GetFilePath("1.docx");
string filePath2 = GetFilePath("2.docx");

Table table1 = GetLastTable(filePath1);

static string[] GetStringsFromParagraph(string filePath1) =>
    GetFirstParagraph(filePath1).InnerText.Split(' ', StringSplitOptions.RemoveEmptyEntries);

string monthName = GetStringsFromParagraph(filePath1)[^3];

int month = Array.IndexOf(DateTimeFormatInfo.CurrentInfo.MonthNames, monthName) + 1;

int year = int.Parse(GetStringsFromParagraph(filePath1)[^2]);

string lastDayMonth = table1.Elements<TableRow>().ElementAt(0).Elements<TableCell>().Last().InnerText;

int daysInMonth = DateTime.DaysInMonth(year, month);

if (int.Parse(lastDayMonth) != daysInMonth)
    throw new Exception("last days");

var incorrectWeekends = table1.Elements<TableRow>().ElementAt(0).Elements<TableCell>().
                            Skip(3).
                            Where(cells => cells.Elements<TableCellProperties>().ElementAt(0).Shading is not null &&
                                new DateOnly(year, month, int.Parse(cells.InnerText)).DayOfWeek is not (DayOfWeek.Saturday or DayOfWeek.Sunday));

if (incorrectWeekends.Any())
    throw new Exception("incorrect weekends");


var schedule = new List<List<string>>();

foreach (var row in table1.Elements<TableRow>())
{
    var name = new List<string>() { row.ElementAt(2).InnerText };
    var days = row.Elements<TableCell>().Skip(3).Select(c => c.InnerText);

    name.AddRange(days);
    schedule.Add(name);
}

for (int i = 1; i < schedule.Count; i++)
{
    if (schedule.Count(r => r[i] is "Х" or "X") is not (2 or 3))
        throw new Exception($"{i}");
}

var weekendsWithEights = table1.Elements<TableRow>().ElementAt(3).Elements<TableCell>().
                            Skip(3).
                            Where(cells => cells.Elements<TableCellProperties>().ElementAt(0).Shading is not null && cells.InnerText == "8");


if (weekendsWithEights.Any())
    throw new Exception("weekend eights");

Table table2 = GetLastTable(filePath2);

var names2 = table2.Elements<TableRow>().Where(rows => rows.Elements<TableCell>().Last().InnerText is "Х" or "X").
                                        Select(rows => rows.Elements<TableCell>().ElementAt(1).InnerText);

var names1 = table1.Elements<TableRow>().Where(rows => rows.Elements<TableCell>().ElementAt(3).InnerText is "Х" or "X" or "8").
                                        Select(rows => rows.Elements<TableCell>().ElementAt(1).InnerText);

if (names1.Intersect(names2).Any())
    throw new Exception("first days");

static Body GetBody(string filePath)
{
    using WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filePath, false);

    return wordprocessingDocument.MainDocumentPart.Document.Body;
}

static Table GetLastTable(string filePath) => GetElements<Table>(filePath).Last();

static Paragraph GetFirstParagraph(string filePath) => GetElements<Paragraph>(filePath).First();

static IEnumerable<T> GetElements<T>(string filePath) where T : OpenXmlElement => GetBody(filePath).Elements<T>();

static string GetFilePath(string fileName)
{
    string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
    string file = Path.Combine(currentDirectory, @"../../../static/" + fileName);
    string filePath = Path.GetFullPath(file);

    return filePath;
}