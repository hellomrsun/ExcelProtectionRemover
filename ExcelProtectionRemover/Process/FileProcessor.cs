using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelProtectionRemover.Process
{
    public class FileProcessor
    {
        public async Task<string> Process(string fileFullPath)
        {
            var convertedFileFullPath = CreateConvertedFileFullPath(fileFullPath);

            CopyFileToConverted(fileFullPath, convertedFileFullPath);

            var zipFileFullPath = RenameFileToZipFile(convertedFileFullPath);

            var zipFolderName = ExtractZipFile(zipFileFullPath);

            RemoveSheetProtection(zipFolderName, zipFileFullPath);

            RemoveAllFilesAndFoldersFromFolder(zipFolderName);

            convertedFileFullPath = Path.ChangeExtension(zipFileFullPath, ".xlsx");

            RemoveFile(convertedFileFullPath);

            File.Move(zipFileFullPath, convertedFileFullPath);

            return convertedFileFullPath;
        }

        private static void RemoveSheetProtection(string zipFolderName, string zipFileFullPath)
        {
            var sheet1FilePath = Path.Combine(zipFolderName, @"xl\worksheets", "sheet1.xml");
            var allContents = File.ReadAllText(sheet1FilePath);

            var regex = new Regex("<sheetProtection[a-zA-Z0-9/ -=\"]+\\/>");
            var newAllContents = regex.Replace(allContents, "");

            File.WriteAllText(sheet1FilePath, newAllContents);

            ZipFile.CreateFromDirectory(zipFolderName, zipFileFullPath);
        }

        private static string ExtractZipFile(string zipFileFullPath)
        {
            var zipFile = new FileInfo(zipFileFullPath);
            var zipFileName = Path.GetFileNameWithoutExtension(zipFileFullPath);
            var zipFolderName = Path.Combine(zipFile.DirectoryName, zipFileName);
            RemoveAllFilesAndFoldersFromFolder(zipFolderName);

            ZipFile.ExtractToDirectory(zipFileFullPath, zipFolderName);
            File.Delete(zipFileFullPath);
            return zipFolderName;
        }

        private static void RemoveAllFilesAndFoldersFromFolder(string zipFolderName)
        {
            if (Directory.Exists(zipFolderName))
            {
                var directory = new DirectoryInfo(zipFolderName);
                foreach (var file in directory.GetFiles()) file.Delete();
                foreach (var subDirectory in directory.GetDirectories()) subDirectory.Delete(true);
                Directory.Delete(zipFolderName);
            }
        }

        private static string RenameFileToZipFile(string convertedFileFullPath)
        {
            var zipFileFullPath = Path.ChangeExtension(convertedFileFullPath, ".zip");

            RemoveFile(zipFileFullPath);

            File.Move(convertedFileFullPath, zipFileFullPath);
            return zipFileFullPath;
        }

        private static void RemoveFile(string zipFileFullPath)
        {
            if (File.Exists(zipFileFullPath))
            {
                File.Delete(zipFileFullPath);
            }
        }

        private static string CreateConvertedFileFullPath(string fileFullPath)
        {
            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileFullPath);
            var newFileName = fileNameWithoutExtension + "_Converted";
            var excelFile = new FileInfo(fileFullPath);
            return Path.Combine(excelFile.DirectoryName, newFileName + excelFile.Extension);
        }

        private static void CopyFileToConverted(string oldFileFullPath, string newFileFullPath)
        {
            if (File.Exists(newFileFullPath))
            {
                File.Delete(newFileFullPath);
            }

            File.Copy(oldFileFullPath, newFileFullPath);
        }
    }
}
