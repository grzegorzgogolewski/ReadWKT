using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using CommandLine;
using License;
using OfficeOpenXml;
using ReadWKT.Tools;
using ReadWKT.VB;
using WKT_Tools;
using System.Threading;

namespace ReadWKT
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            ConsoleColor defaultColor = Console.ForegroundColor;

            MyLicense license = LicenseHandler.ReadLicenseFile(out LicenseStatus licStatus, out string validationMsg);
            
            switch (licStatus)
            {
                case LicenseStatus.Undefined:
                    
                    Console.ForegroundColor = ConsoleColor.Red; 
                    Console.WriteLine("Brak pliku z licencją!!!\n");
                    
                    Console.ForegroundColor = defaultColor;

                    Assembly assembly = Assembly.GetExecutingAssembly();
                    FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);

                    Console.WriteLine("Identyfikator komputera: " + LicenseHandler.GenerateUid(fvi.ProductName) + '\n');

                    Console.ReadKey(false);
                    Environment.Exit(0);
                    break;

                case LicenseStatus.Valid:
                    
                    Console.WriteLine("Właściciel licencji:\n");
                    Console.WriteLine(license.LicenseOwner + "\n");

                    Console.ForegroundColor = ConsoleColor.Blue; 
                    Console.WriteLine($"Licencja typu: '{license.Type}', ważna do: {license.LicenseEnd}\n");
                    
                    Console.ForegroundColor = defaultColor;

                    Thread.Sleep(1000);
                    break;

                case LicenseStatus.Invalid:
                case LicenseStatus.Cracked:

                    Console.ForegroundColor = ConsoleColor.Red; 
                    Console.WriteLine(validationMsg);
                    
                    Console.ForegroundColor = defaultColor;

                    Console.ReadKey(false);
                    Environment.Exit(0);

                    break;

                case LicenseStatus.Expired:
                   
                    Console.WriteLine("Właściciel licencji:");
                    Console.WriteLine(license.LicenseOwner + "\n");

                    Console.ForegroundColor = ConsoleColor.Red; 
                    Console.WriteLine(validationMsg);
                    
                    Console.ForegroundColor = defaultColor;

                    Console.ReadKey(false);
                    Environment.Exit(0);

                    break;

                default:
                    throw new ArgumentOutOfRangeException();
            }

            // --------------------------------------------------------------------------------------------
            // podłączenie bibiotek SQL Server do obsługi geometrii
            
            SqlServerTypes.Utilities.LoadNativeAssemblies(AppDomain.CurrentDomain.BaseDirectory);
            
            // --------------------------------------------------------------------------------------------
            // folder startowy dla danych analizowanych przez aplikację

            string startupPath = string.Empty;

            void RunOptions(Options opts)
            {
                startupPath = opts.StarupPath;
            }

            void HandleParseError(IEnumerable<Error> errs)
            {
                Console.ReadKey(false);
                Environment.Exit(0);
            }

            Parser.Default.ParseArguments<Options>(args).WithParsed(RunOptions).WithNotParsed(HandleParseError);

            // --------------------------------------------------------------------------------------------

            Console.WriteLine("Pobieranie listy folerów...");

            // pobierz wszystkie katalogi i podkatalogi
            string[] allDirectories = Directory.GetDirectories(startupPath, "*", SearchOption.AllDirectories);

            Console.WriteLine($"Pobrano {allDirectories.Length} folderów.\n");

            Console.WriteLine("Analizowanie folderów...");

            // sortowanie nazw katalogów zgodnie z sortowaniem naturalnym
            Array.Sort(allDirectories, new NaturalStringComparer());

            // lista "ostatnich" podfolderów - folder końcowy w którym powinny być dane
            List<string> wktDirectories = new List<string>();

            // analiza podkatalogów pod kątem tego czy są ostatnimi folderami w ścieżce
            foreach (string dir in allDirectories)
            {
                DirectoryInfo directory = new DirectoryInfo(dir); // pobierz informacje o danym folderze
                DirectoryInfo[] subdirs = directory.GetDirectories(); // pobierz list podfolderów dla danego folderu

                if (subdirs.Length == 0) wktDirectories.Add(dir);   // jeśli dany katalog nie ma podfolderów dodaj go do listy katalogów z wkt
            }

            Console.WriteLine("Zliczanie plików...");
            FileInfo[] files = new DirectoryInfo(startupPath).GetFiles("*.wkt", SearchOption.AllDirectories);

            Console.WriteLine($"Koniec analizy folderów. Pozostało {wktDirectories.Count} folderów. Plików WKT: {files.Length}\n");

            int filesCounter = 1;

            Timer timer = new Timer(TimerCallback, null, 0, 1000);

            void TimerCallback(Object o) {
                Console.WriteLine($"ID: {filesCounter, 6} / {files.Length}");
            }

            WktFeatures wktFeatures = new WktFeatures();

            foreach (string wktDirectory in wktDirectories)
            {
                List<string> wktFileNames = Directory.GetFiles(wktDirectory, "*.wkt", SearchOption.TopDirectoryOnly).ToList();

                foreach (string wktFileName in wktFileNames)
                {
                    WktFile wktFile = new WktFile(wktFileName);

                    if (!wktFile.IsValid())
                    {
                        Console.WriteLine($"ID: {filesCounter, 6} => {Path.GetFileName(wktFileName), 40}: {wktFile.ValidationStatus}");
                    }
                    
                    wktFeatures.Add(filesCounter++, wktFile);
                }
            }

            timer.Dispose();

            string outputFile = Path.Combine(startupPath, "wkt.xlsm");

            Console.WriteLine($"\nZapisywanie wyników do pliku {outputFile}...");

            File.Delete(outputFile);

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet arkusz = excelPackage.Workbook.Worksheets.Add("WKT");

                arkusz.Cells[1, 1].LoadFromCollection(wktFeatures.Values, true);

                arkusz.Cells["A1:G1"].AutoFilter = true;
                arkusz.View.FreezePanes(2, 1);
                arkusz.Cells.AutoFitColumns(8.43, 40);

                // --------------------------------------------------------------------------------
                // Dodanie kodu makra
                // --------------------------------------------------------------------------------

                excelPackage.Workbook.CreateVBAProject();

                excelPackage.Workbook.VbaProject.Modules.AddModule("mdlMain").Code = VbResource.GetVbText("mdlMain.txt");

                excelPackage.SaveAs(new FileInfo(outputFile));
            }

            Console.WriteLine("\nKoniec!");
            Console.ReadKey(false);
        }
    }
}
