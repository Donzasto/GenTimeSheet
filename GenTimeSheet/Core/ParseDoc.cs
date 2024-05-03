using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;

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

        private Paragraph GetFirstParagraph() =>GetElements<Paragraph>().First();

        internal string[] GetStringsFromParagraph() =>
            GetFirstParagraph().InnerText.Split(' ', StringSplitOptions.RemoveEmptyEntries);
    }
}
