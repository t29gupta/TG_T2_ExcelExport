using System;
using System.IO;
using System.Text.Json;

namespace TG_Reporting
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("TG Assignment 2. Export excel report from json File\n");

            var hotelRatesjsonPath = Path.Combine(AppContext.BaseDirectory, "hotelrates.json");

            // Read json file
            var hotelRates = File.ReadAllText(hotelRatesjsonPath);

            var excelGenerator = new ExcelReportGenerater();

            Console.WriteLine("Generation excel from string...Press enter to continue.");
            Console.ReadLine();
            
            var outPutFromString = excelGenerator.GenerateExcelReport(hotelRates);
            if (string.IsNullOrEmpty(outPutFromString))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error occured while generating excel report.");
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Output from string saved at: {outPutFromString}\n");
            }
            Console.ResetColor();

            Console.WriteLine("Generation excel from stream...Press enter to continue.");
            Console.ReadLine();

            using (var fileStream = File.OpenRead(hotelRatesjsonPath))
            {
                var outPutFromStream = excelGenerator.GenerateExcelReport(fileStream);

                if (string.IsNullOrEmpty(outPutFromStream))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Error occured while generating excel report.");
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"Output from stream saved at: {outPutFromStream }\n");
                }
                Console.ResetColor();
            }

            Console.WriteLine("Generation excel from json object...Press enter to continue.");
            Console.ReadLine();

            var outPutFromJsonObject = excelGenerator.GenerateExcelReport(JsonDocument.Parse(hotelRates));
            if (string.IsNullOrEmpty(outPutFromJsonObject))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error occured while generating excel report.");
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Output from json object saved at: {outPutFromJsonObject }\n");
            }
            Console.ResetColor();

            Console.WriteLine("Generation excel from fileinfo...Press enter to continue.");
            Console.ReadLine();

            var outPutFromFileInfo = excelGenerator.GenerateExcelReport(new FileInfo(hotelRatesjsonPath));
            if (string.IsNullOrEmpty(outPutFromFileInfo))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error occured while generating excel report.");
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Output from Fileinfo saved at: {outPutFromFileInfo}\n");
            }

            Console.ResetColor();
            Console.WriteLine("Press exnter to exit...");
            Console.ReadLine();
        }
    }
}
