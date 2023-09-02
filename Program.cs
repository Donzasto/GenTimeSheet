using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

string filePath1 = GetFilePath("1.docx");
string filePath2 = GetFilePath("2.docx");

Table? table1 = GetTable(filePath1);

var schedule = new List<List<string>>();

foreach (var row in table1.Elements<TableRow>())
{
    var name = new List<string>() { row.ElementAt(2).InnerText };
    var days = row.Elements<TableCell>().Skip(3).Select(c => c.InnerText);

    name.AddRange(days);
    schedule.Add(name);
}

for (int i = 1; i < schedule.Count; i++)
    if (schedule.Count(r => r[i] is "Х" or "X") is not 2 or 3)
        throw new Exception($"{i}");


if (!schedule[3].Contains("8")) throw new Exception("eights");

Table? table2 = GetTable(filePath2);

var names2 = table2.Elements<TableRow>().Where(w => w.Elements<TableCell>().Last().InnerText is "Х" or "X")
                                       .Select(s => s.Elements<TableCell>().ElementAt(1).InnerText);

var names1 = table1.Elements<TableRow>().Where(w => w.Elements<TableCell>().ElementAt(3).InnerText is "Х" or "X")
                                        .Select(s => s.Elements<TableCell>().ElementAt(1).InnerText);

if (names1?.Intersect(names2).Count() != 0) throw new Exception("first");

static Table? GetTable(string filePath)
{
    using WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filePath, false);

    return wordprocessingDocument.MainDocumentPart?.Document.Body?.Elements<Table>().Last();
}

static string GetFilePath(string fileName)
{
    string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
    string file = Path.Combine(currentDirectory, @"../../../static/" + fileName);
    string filePath = Path.GetFullPath(file);

    return filePath;
}