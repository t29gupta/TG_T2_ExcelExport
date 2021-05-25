using NUnit.Framework;
using System;
using System.IO;
using System.Text.Json;

namespace TG_Reporting.Tests
{
    public class Tests
    {
        private string mockFilePath;
        private ExcelReportGenerater excelGenerator;

        [SetUp]
        public void Setup()
        {
            mockFilePath = Path.Combine(AppContext.BaseDirectory, "mockHotelrates.json");
            excelGenerator = new ExcelReportGenerater();
        }

        [Test]
        public void GenerateExcelFromStringContent_ReturnsPath()
        {
            var hotelRates = File.ReadAllText(mockFilePath);
            var outputFilePath = excelGenerator.GenerateExcelReport(hotelRates);

            Assert.IsFalse(string.IsNullOrEmpty(outputFilePath));
            Assert.IsTrue(File.Exists(outputFilePath));
        }

        [Test]
        public void GenerateExcelFromContentStream_ReturnsPath()
        {
            using (var fileStream = File.OpenRead(mockFilePath))
            {
                var outputFilePath = excelGenerator.GenerateExcelReport(fileStream);

                Assert.IsFalse(string.IsNullOrEmpty(outputFilePath));
                Assert.IsTrue(File.Exists(outputFilePath));
            }
        }

        [Test]
        public void GenerateExcelFromJsonObject_ReturnsPath()
        {
            var hotelRates = File.ReadAllText(mockFilePath);
            using (var jsonObject = JsonDocument.Parse(hotelRates))
            {
                var outputFilePath = excelGenerator.GenerateExcelReport(jsonObject);

                Assert.IsFalse(string.IsNullOrEmpty(outputFilePath));
                Assert.IsTrue(File.Exists(outputFilePath));
            }
        }

        [Test]
        public void GenerateExcelFromFileInfo_ReturnsPath()
        {
            var hotelRatesFileInfo = new FileInfo(mockFilePath);
            var outputFilePath = excelGenerator.GenerateExcelReport(hotelRatesFileInfo);

            Assert.IsFalse(string.IsNullOrEmpty(outputFilePath));
            Assert.IsTrue(File.Exists(outputFilePath));
        }
    }
}