using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GenTimeSheet.Core
{
    internal class ParseDoc
    {
        private readonly string _filePath;

        internal ParseDoc(string fileName)
        {
            _filePath = FileHandler.GetFilePath(fileName);
        }

        internal Table GetTableFromEnd(int index) => GetElements<Table>().ElementAt(^index);

        internal string[] GetStringsFromParagraph() =>
            GetFirstParagraph().InnerText.Split(' ', StringSplitOptions.RemoveEmptyEntries);

        private IEnumerable<T> GetElements<T>() where T : OpenXmlElement =>
            GetBody().Elements<T>();

        private Body GetBody()
        {
            try
            {
                using var wordprocessingDocument = WordprocessingDocument.Open(_filePath, false);

                return wordprocessingDocument.MainDocumentPart.Document.Body;
            }
            catch
            {
                throw;
            }
        }

        private Paragraph GetFirstParagraph() => GetElements<Paragraph>().First();

        internal IEnumerable<string> GetNamesWorkedLastDayMonth(Table previousMonthTable) => previousMonthTable.Elements<TableRow>().
            Where(rows => rows.Elements<TableCell>().Last().InnerText.EqualsOneOf(Constants.RU_X, Constants.EN_X)).
            Select(rows => Regex.Replace(rows.Elements<TableCell>().ElementAt(1).InnerText, @"\s+",
                string.Empty)).ToArray();
    }
}
