using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using NPOI.XSSF.UserModel;  // Für das .xlsx-Format
using NPOI.SS.UserModel;   // Gemeinsame Schnittstellen für xls und xlsx
using ClosedXML.Excel;     // Für das .xlsx-Format


namespace FileWork
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //PrintAllEnvironmentPaths();
            //PrintAllDrives();
            //PrintAllSystemPaths();
            //CheckIfFileExists();
            //CheckIfPathExists();
            //OpenExplorer(Environment.CurrentDirectory);
            //ReadWriteFile();
            //CreateExcelFileWithNPOI($"{Environment.CurrentDirectory}\\Datei.xlsx");
            //ReadExcelFileWithNPOI($"{Environment.CurrentDirectory}\\Datei.xlsx");
            //ReadExcelFileWithClosedXML($"{Environment.CurrentDirectory}\\Datei.xlsx");
        }



        public static void ReadExcelFileWithClosedXML(string filePath)
        {
            Console.WriteLine();

            using (var workbook = new XLWorkbook(filePath)) //using stellt sicher das die Datei nach dem Verlassen des Blocks geschlossen wird
            {
                var worksheet = workbook.Worksheets.First(); // Nimmt das erste Arbeitsblatt
                              
                for (int rowNumber = 1; rowNumber <= worksheet.LastRowUsed().RowNumber(); rowNumber++)
                {
                    var row = worksheet.Row(rowNumber);
                    int lastColumn = row.LastCellUsed().Address.ColumnNumber; // Letzte Spalte mit einem Wert

                    for (int colNumber = 1; colNumber <= lastColumn; colNumber++)
                    {
                        var cellValue = row.Cell(colNumber).Value.ToString();
                        Console.Write($"{cellValue.PadRight(20)}|");
                    }
                    Console.WriteLine();
                }            
            }
        }

        static void PrintAllEnvironmentPaths()
        {
            Console.WriteLine();
            Console.WriteLine("System Paths:");
            Console.WriteLine($"Current Directory: {Environment.CurrentDirectory}");
            Console.WriteLine($"System Directory: {Environment.SystemDirectory}");
            Console.WriteLine($"User Domain Name: {Environment.UserDomainName}");
            Console.WriteLine($"User Name: {Environment.UserName}");
            Console.WriteLine($"Machine Name: {Environment.MachineName}");
            Console.WriteLine($"OS Version: {Environment.OSVersion}");
            Console.WriteLine($"Processor Count: {Environment.ProcessorCount}");
            Console.WriteLine($"Version: {Environment.Version}");
            Console.WriteLine($"Working Set: {Environment.WorkingSet}");
        }

        static void PrintAllDrives()
        {
            Console.WriteLine("Drives:");
            foreach (var drive in DriveInfo.GetDrives())
            {
                try
                {
                    Console.WriteLine();
                    Console.WriteLine($"Drive: {drive.Name}");
                    Console.WriteLine($"Drive Type: {drive.DriveType}");
                    Console.WriteLine($"Drive Format: {drive.DriveFormat}");
                    Console.WriteLine($"Drive Volume Label: {drive.VolumeLabel}");
                    Console.WriteLine($"Drive Available Free Space: {drive.AvailableFreeSpace}");
                    Console.WriteLine($"Drive Total Free Space: {drive.TotalFreeSpace}");
                    Console.WriteLine($"Drive Total Size: {drive.TotalSize}");
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }

        static void PrintAllSystemPaths()
        {

            Console.WriteLine();

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            Console.WriteLine($"Desktop Path: {desktopPath}");

            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            Console.WriteLine($"Documents Path: {documentsPath}");

            string programFilesPath = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            Console.WriteLine($"Program Files Path: {programFilesPath}");

            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            Console.WriteLine($"AppData Path: {appDataPath}");

            string systemPath = Environment.GetFolderPath(Environment.SpecialFolder.System);
            Console.WriteLine($"System Directory Path: {systemPath}");

            string picturesPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonPictures);
            Console.WriteLine($"Pictures Path: {picturesPath}");

            //Or you can do it in this way
            Console.WriteLine($"AdminTools Path: {Environment.GetFolderPath(Environment.SpecialFolder.AdminTools)}");
            Console.WriteLine($"CommonProgramFilesX86 Path: {Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFilesX86)}");
            Console.WriteLine($"ProgramFiles Path: {Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)}");
            Console.WriteLine($"Windows Path: {Environment.GetFolderPath(Environment.SpecialFolder.Windows)}");
            Console.WriteLine($"Favorites Path: {Environment.GetFolderPath(Environment.SpecialFolder.Favorites)}");
        }

        static void CheckIfFileExists()
        {
            string path = $"{Environment.CurrentDirectory}\\calc.exe";
            Console.WriteLine(path);

            if (File.Exists(path))
            {
                Console.WriteLine("File exists");
            }
            else
            {
                Console.WriteLine("File does not exist");
            }
        }

        static void CheckIfPathExists()
        {
            string path = @"C:\Windows\System32\";
            Console.WriteLine(path);

            if (Directory.Exists(path))
            {
                Console.WriteLine("Path exists");
            }
            else
            {
                Console.WriteLine("Path does not exist");
            }
        }

        static void OpenExplorer(string path)
        {
            // Prüfen, ob der Pfad existiert
            if (System.IO.Directory.Exists(path))
            {
                // Startet den Windows Explorer im angegebenen Pfad
                Process.Start("explorer.exe", path);
                Console.WriteLine($"Explorer geöffnet im Pfad: {path}");
            }
            else
            {
                Console.WriteLine($"Der angegebene Pfad existiert nicht: {path}");
            }
        }

        static void ReadWriteFile()
        {
            string path = $"{Environment.CurrentDirectory}\\";
            string file = "test.txt";
            Console.WriteLine(path);

            //Write to file
            File.WriteAllText(path + file, "Hello World");

            //Read from file
            string content = File.ReadAllText(path + file);
            Console.WriteLine(content);

            OpenExplorer(path);
        }

        static void CreateExcelFileWithNPOI(string filePath)
        {
            Console.WriteLine();

            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            XSSFWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet1");

            // Eine Zeile und Zellen hinzufügen
            IRow row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue("Name");
            row.CreateCell(1).SetCellValue("Alter");
            row.CreateCell(2).SetCellValue("Ort");

            // Daten in die nächste Zeile schreiben
            IRow dataRow = sheet.CreateRow(1);
            dataRow.CreateCell(0).SetCellValue("Max Mustermann");
            dataRow.CreateCell(1).SetCellValue(29);
            dataRow.CreateCell(2).SetCellValue("Berlin");

            // Datei speichern
            using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(file);
            }

            Console.WriteLine("Excel-Datei erfolgreich erstellt!");
        }

        static void ReadExcelFileWithNPOI(string filePath)
        {
            Console.WriteLine();            

            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                XSSFWorkbook workbook = new XSSFWorkbook(file);
                ISheet sheet = workbook.GetSheetAt(0); // Erstes Arbeitsblatt

                for (int row = 0; row <= sheet.LastRowNum; row++) // Zeilen durchlaufen
                {
                    IRow currentRow = sheet.GetRow(row);
                    if (currentRow != null)
                    {
                        for (int col = 0; col < currentRow.LastCellNum; col++) // Spalten durchlaufen
                        {
                            ICell cell = currentRow.GetCell(col);
                            Console.Write(cell?.ToString().PadRight(20) + "|");
                        }
                        Console.WriteLine();
                    }
                }
            }
        }
    }
}