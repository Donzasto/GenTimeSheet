using System;
using System.IO;

namespace GenTimeSheet.Core;

public static class FileHandler
{
    public static string GetFilePath(string fileName)
    {
        string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
        string file = Path.Combine(currentDirectory, @"../../../../GenTimeSheet/files/" + fileName);
        string filePath = Path.GetFullPath(file);

        return filePath;
    }
}
