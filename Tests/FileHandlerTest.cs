using GenTimeSheet.Core;

namespace Tests
{
    public class FileHandlerTest
    {
        [Fact]
        public void CheckFileExist()
        {
            string filePath = FileHandler.GetFilePath("testFiles/1.xlsx");

            Assert.True(File.Exists(filePath));
        }

        [Fact]
        public void CheckFileNotExist()
        {
            string filePath = FileHandler.GetFilePath("testFiles/2.xlsx");

            Assert.False(File.Exists(filePath));
        }
    }
}