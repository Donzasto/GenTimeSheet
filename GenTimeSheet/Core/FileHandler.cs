using System;
using System.IO;

namespace GenTimeSheet.Core
{
    internal static class FileHandler
    {
        internal static string GetFilePath(string fileName)
        {
            string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string file = Path.Combine(currentDirectory, @"../../../../GenTimeSheet/static/" + fileName);
            string filePath = Path.GetFullPath(file);

            return filePath;
        }
    }
}
