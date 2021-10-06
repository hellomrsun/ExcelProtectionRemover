using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;

namespace ExcelProtectionRemover.Process
{
    public class ExcelFileProcessor
    {
        public string Process(string fileFullPath)
        {
            var convertedFileFullPath = CreateConvertedFileFullPath(fileFullPath);

            CopyFileToConverted(fileFullPath, convertedFileFullPath);

            var zipFileFullPath = RenameFileToZipFile(convertedFileFullPath);

            var zipFolderName = ExtractZipFile(zipFileFullPath);

            RemoveExcelSheetsProtection(zipFolderName, zipFileFullPath);

            RemoveAllFilesAndFoldersFromFolder(zipFolderName);

            convertedFileFullPath = Path.ChangeExtension(zipFileFullPath, ".xlsx");

            TryRemoveFile(convertedFileFullPath);

            File.Move(zipFileFullPath, convertedFileFullPath);

            return convertedFileFullPath;
        }

        private static void RemoveExcelSheetsProtection(string zipFolderName, string zipFileFullPath)
        {
            var folder = Path.Combine(zipFolderName, @"xl\worksheets");
            var files = Directory.GetFiles(folder);

            foreach (var filePath in files)
            {
                var allContents = File.ReadAllText(filePath);
                var regex = new Regex("<sheetProtection[a-zA-Z0-9/ -=\"]+\\/>");
                var newAllContents = regex.Replace(allContents, "");
                File.WriteAllText(filePath, newAllContents);
            }
            
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
            if (!Directory.Exists(zipFolderName)) return;
            var directory = new DirectoryInfo(zipFolderName);
            foreach (var file in directory.GetFiles()) file.Delete();
            foreach (var subDirectory in directory.GetDirectories()) subDirectory.Delete(true);
            Directory.Delete(zipFolderName);
        }

        private static string RenameFileToZipFile(string convertedFileFullPath)
        {
            var zipFileFullPath = Path.ChangeExtension(convertedFileFullPath, ".zip");

            TryRemoveFile(zipFileFullPath);

            File.Move(convertedFileFullPath, zipFileFullPath);
            return zipFileFullPath;
        }

        private static void TryRemoveFile(string zipFileFullPath)
        {
            if (File.Exists(zipFileFullPath))
            {
                File.Delete(zipFileFullPath);
            }
        }

        private static string CreateConvertedFileFullPath(string fileFullPath)
        {
            var newFileName = Path.GetFileNameWithoutExtension(fileFullPath) + "_ProtectionRemoved";
            var excelFile = new FileInfo(fileFullPath);
            return Path.Combine(excelFile.DirectoryName, newFileName + excelFile.Extension);
        }

        private static void CopyFileToConverted(string oldFileFullPath, string newFileFullPath)
        {
            TryRemoveFile(newFileFullPath);

            File.Copy(oldFileFullPath, newFileFullPath);
        }
    }
}
