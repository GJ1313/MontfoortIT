using Microsoft.VisualStudio.TestTools.UnitTesting;
using MontfoortIT.Library.Streams.FileConvertors;
using MontfoortIT.Office.Excel.Csv;
using MontfoortIT.Office.Excel.Standard.Templates;
using MontfoortIT.Office.Excel.Templates;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace MontfoortIT.Office.Excel.Standard.Test.Csv
{
    [TestClass]
    public class WriteCsvTest
    {
        [TestMethod]
        public async Task WritePdfTest()
        {
            List<TestClass> items = new List<TestClass>()
            {
                new TestClass(){ Text="Rij 1", Number=1},
                new TestClass(){Text = "rij 2", Number=2}
            };

            FuncTemplateList<TestClass> template = new FuncTemplateList<TestClass>();
            template.Add("Text", f => f.Text);
            template.Add("Number", f => f.Number);

            CsvSheet sheet = new CsvSheet();
            sheet.ColumnTemplate = template;

            MemoryStream stream = new MemoryStream();
            await sheet.WriteFromObjectsAsync(items, stream, new Options(), Encoding.ASCII);

            stream.Position = 0;
            string result = Encoding.ASCII.GetString(stream.ToArray());

            string[] parts = result.Split('\n');
            Assert.AreEqual(4, parts.Length);

            StringAssert.Contains("Rij 1,1", parts[1].Trim());
        }

        [TestMethod]
        public async Task ConvertCsvToExcelTest()
        {
            // Assemble
            var csvStream = File.OpenRead(@"Files/Belgie.csv");            
            using Application application = new Application();
            var newSheet = application.Workbook.Sheets.Create("New sheet");

            CommaSeperatedToXml commaSeperator = new CommaSeperatedToXml();


            // Act

            await commaSeperator.ConvertAsync(csvStream, new CsvToExcelWriter(application, newSheet));
            
            //await newSheet.FillFromCsvAsync(csvStream, new Options());

            // Assert
            Assert.AreEqual("110", newSheet.Cells[1, 1].Text);

        }

    }

    internal class TestClass
    {
        public string Text { get; set; }

        public int Number { get; set; }
    }
}
