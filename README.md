# FileWork
This is a simple program that interacts with the file system using .NET Core 8.x.

## Get System Path
	PrintAllSystemPaths();

## Get Environment Path
	PrintAllEnvironmentPaths();

## Get Drive Path	
	PrintAllDrives();

## Check if File Exists
	CheckIfFileExists();

## Check if Directory Exists
	CheckIfDirectoryExists();

## Open Windows Explorer 	
	OpenExplorer(string path);

## Read/Write File
	ReadWriteFile();

## Excel File
Um eine .xlsx-Datei (das neuere Excel-Format) in C# zu lesen und zu schreiben, 
können Sie die NPOI-Bibliothek oder die ClosedXML verwenden, die neben .xls auch das .xlsx-Format unterstützt.

- Um die ClosedXML-Bibliothek in einem .NET-Projekt zu installieren, können Sie den **NuGet Package Manager** 	
in Visual Studio oder die .NET CLI verwenden oder in *.csproj Datei 
<PackageReference Include="ClosedXML" Version="0.97.0" /> hinzufügen.
    
- Um die NPOI-Bibliothek in einem .NET-Projekt zu installieren, können Sie den **NuGet Package Manager** 
in Visual Studio oder die .NET CLI verwenden oder in *.csproj Datei 
<PackageReference Include="NPOI" Version="2.7.1" /> hinzufügen.

### Installation über den NuGet Package Manager in Visual Studio
Falls Sie Visual Studio verwenden, können Sie das Paket folgendermaßen hinzufügen:

1. Projekt öffnen: Öffnen Sie Ihr Projekt in Visual Studio.
2. NuGet Package Manager öffnen:
 * Klicken Sie mit der rechten Maustaste auf Ihr Projekt im Solution Explorer.
 * Wählen Sie Manage NuGet Packages (NuGet-Pakete verwalten).
3. *DAS* Paket suchen:
 * Wechseln Sie zum Tab Browse (Durchsuchen).
 * Geben Sie den „NAMEN“ in das Suchfeld ein.
 * Wählen Sie das Paket aus der Liste der verfügbaren Pakete.
4. Paket installieren:
 * Klicken Sie auf Install (Installieren), um das Paket in Ihrem Projekt zu installieren.
5. Paketinstallation überprüfen:
 * Nach der Installation sollten Sie das Paket unter dem Tab Installed (Installierte Pakete) sehen.

Bei beiden beispielen **Schreiben** **Lesen** benötiten Sie die entsprechende Namespaces:
```csharp
using ClosedXML.Excel;     // Für das .xlsx-Format
```

```csharp
using NPOI.XSSF.UserModel;  // Für das .xlsx-Format
using NPOI.SS.UserModel;   // Gemeinsame Schnittstellen für xls und xlsx
```


### Write XLSX (Excel File)
	CreateExcelFileWithNPOI($"{Environment.CurrentDirectory}\\Datei.xlsx");            

#### Erklärung des Codes:

- HSSFWorkbook: Erstellen Sie ein neues Arbeitsbuch für das .xls-Format.
- CreateSheet: Erstellt ein neues Arbeitsblatt.
- CreateRow und CreateCell: Diese Methoden fügen der Arbeitsmappe neue Zeilen und Zellen hinzu.
- SetCellValue: Legt den Wert einer Zelle fest, sei es Text, Zahl oder Datum.
- Write: Speichert die Arbeitsmappe in der angegebenen Datei.

### Read XLSX (Excel File)
	ReadExcelFileWithClosedXML($"{Environment.CurrentDirectory}\\Datei.xlsx");-
	ReadExcelFileWithNPOI($"{Environment.CurrentDirectory}\\Datei.xlsx");

#### Erklärung des Codes:
HSSFWorkbook: Dies ist der Hauptarbeitsbuchtyp für .xls-Dateien 
ISheet: Stellt ein Arbeitsblatt dar. Sie können es durch Angabe des Index (z. B. 0 für das erste Blatt) abrufen.
IRow und ICell: Diese Objekte repräsentieren Zeilen und Zellen, die Sie durchlaufen und deren Inhalte abrufen können.





Read/Write JSON Files
Read/Write XML Files
Read/Write Binary Files
Read/Write to Network Shares


Copy File
Move File
Get File Properties
Get File Permissions
Set File Permissions
Write to File
Append to File
Delete File
Create Directory
Delete Directory
List Files in Directory

Monitor File Changes (using FileSystemWatcher)
Compress/Decompress Files (e.g., Zip)
Encrypt/Decrypt Files
Create Temporary Files/Directories

Handle File Exceptions and Errors
Search for Files and Directories (using patterns and filters)

Get File Size
Get Directory Size
Compare Files (e.g., byte-by-byte comparison)
Create Symbolic Links or Hard Links
Work with File Streams (e.g., BufferedStream, MemoryStream)

Create and Manage Shortcuts
Handle Large Files (e.g., with memory-mapped files)
Get Disk Information (e.g., available space, total space)
Work with Hidden Files and Directories
Set File Attributes (e.g., Read-Only, Hidden)
Get/Set File Creation, Access, and Modification Dates
Check File Lock Status
Restore Files from Recycle Bin
Implement File Versioning (e.g., maintaining multiple versions of a file)
Work with Read-Only Filesystems
Schedule File Operations (e.g., using Task Scheduler or background services)
Log File Activities (e.g., access, modification, deletion)
Work with Special Folders (e.g., AppData, Program Files)
Integrate with Cloud Storage APIs (e.g., Azure Blob Storage, AWS S3)
Create/Extract TAR/GZ Archives
Work with File Metadata (e.g., EXIF data for images)
Work with Resource Files (e.g., embedded resources in assemblies)
Integrate File Operations with ASP.NET Core Web APIs
Create Custom File Formats (e.g., serialization)
Asynchronous File Operations (e.g., async/await for I/O operations)
Implement File Integrity Checks (e.g., using hash algorithms)
Work with External Storage Devices (e.g., USB drives)
Manage File Locking and Concurrency (e.g., using Mutex)
Process Large Files in Chunks
Implement Retry Logic for File Operations
