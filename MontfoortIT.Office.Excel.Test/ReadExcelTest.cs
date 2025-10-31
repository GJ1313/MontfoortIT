using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MontfoortIT.Library.Extensions;
using MontfoortIT.Office.Excel.Templates;

namespace MontfoortIT.Office.Excel.Test
{
    [TestClass]
    public class ReadExcelTest
    {
        private readonly Mutex _testMutex = new Mutex();


        [TestMethod]
        public void ReadSharedStringTest()
        {
            // Assemble
            using (Application application = new Application())
            {

                // Act
                application.ReadFile(@"Files/Clusterindeling_ per_02_september 2013.xlsx");

                // Assert
                Assert.AreEqual(3, application.Workbook.Sheets.Count);
                Sheet sheet1 = application.Workbook.Sheets[0];
                Assert.AreEqual("Rayondirecteur", sheet1.Cells[0, 0].Text);
                Assert.AreEqual("Rayondirecteur", sheet1.Cells["A1"].Text);
            }
        }

        [TestMethod]
        public void ReadNumberAndDateTest()
        {
            _testMutex.WaitOne();
            // Assemble
            using (Application application = new Application())
            {
                // Act
                application.ReadFile(@"DataTest/20131017_Mid-Day_DPA.xlsx");
                
                // Assert
                Assert.AreEqual(3, application.Workbook.Sheets.Count);
                Sheet sheet1 = application.Workbook.Sheets[0];
                Assert.AreEqual("61.7", sheet1.Cells[1, 1].Text);
                Assert.AreEqual("41579", sheet1.Cells[1, 0].Text);
            }
            _testMutex.ReleaseMutex();
        }

        [TestMethod]
        public void ReadFileWithoutErrorTest()
        {
            // Assemble
            _testMutex.WaitOne();
            using (Application application = new Application())
            {                
                // Act
                application.ReadFile(@"DataTest/20131008_Mid-Day_DPA.xlsx");

                
                // Assert
                Assert.AreEqual(3, application.Workbook.Sheets.Count);
                Sheet sheet1 = application.Workbook.Sheets[0];
                Assert.AreEqual("62.24", sheet1.Cells[1, 1].Text);
                Assert.AreEqual("62.24", sheet1.Cells[1, "b"].Text);
                Assert.AreEqual("41579", sheet1.Cells[1, 0].Text);
            }
            _testMutex.ReleaseMutex();
        }

        [TestMethod]
        public void ReadFileWithoutStrangeDateTest()
        {
            // Assemble
            _testMutex.WaitOne();
            using (Application application = new Application())
            {
                // Act
                application.ReadFile(@"DataTest/20131008_Mid-Day_DPB.xlsx");

                
                // Assert
                Assert.AreEqual(3, application.Workbook.Sheets.Count);
                Sheet sheet1 = application.Workbook.Sheets[0];
                Assert.IsTrue(sheet1.Cells["A81"].Text.StartsWith("41555."));
            }
            _testMutex.ReleaseMutex();
        }

        [TestMethod]
        public void TestColumCount()
        {
            _testMutex.WaitOne();
            // Assemble
            using (Application application = new Application())
            {                
                // Act
                application.ReadFile(@"DataTest/20131008_Mid-Day_DPB.xlsx");

                
                Sheet sheet1 = application.Workbook.Sheets[0];

                Cell cell = sheet1.Cells["A5"];
                Assert.AreEqual(0, cell.Column);
                Assert.AreEqual(4, cell.Row);

                cell = sheet1.Cells["Z5"];
                Assert.AreEqual(25, cell.Column);
                Assert.AreEqual(4, cell.Row);

                cell = sheet1.Cells["AA4"];
                Assert.AreEqual(26, cell.Column);
                Assert.AreEqual(3, cell.Row);

                cell = sheet1.Cells["AZ4"];
                Assert.AreEqual(51, cell.Column);
                Assert.AreEqual(3, cell.Row);
            }
            _testMutex.ReleaseMutex();
        }
               

        [TestMethod]
        public async Task UseExcelAsTemplate()
        {
            // Assemble
            using Application application = new Application();

            // Act
            application.ReadFile(@"Files/ScrapeITTemplate.xlsx");
            application.Workbook.Sheets[0].Cells[1, 0].Text = "Kolom 1";
            application.Workbook.Sheets[0].Cells[2, 0].Text = "Rij 1";
            application.Workbook.Sheets[0].Cells[0, 2].Text = "30-10-2015\nGertjan Trading";

            var newSheet = application.Workbook.Sheets.Create("New sheet");
            newSheet.Cells[1, 0].Text = "Sheet 2";


            await application.WriteAsTemplateToAsync(@"C:\Temp\Test2.xlsx");
        }


        [TestMethod]
        public async Task UseExcelAsTemplatev2()
        {
            // Assemble
            using Application application = new Application((n, header) =>
            {
                if (header)
                    return "9";
                return ((int)n).ToString();
            });

            // Act
            application.ReadFile(@"Files/ScrapeITTemplate.xlsx");
            var sheet = application.Workbook.Sheets[0];
            sheet.Cells[0, 2].Text = $"{DateTime.Today:dd-MM-yyyy}\nTestCustomer";

            sheet.ColumnTemplate = new List<ColumnTemplate>()
            {
                new FuncColumnTemplate<TestObject>("Test string", t=>t.Column1),
                new FuncColumnTemplate<TestObject>("Test int", t=>t.Column2),
                new FuncColumnTemplate<TestObject>("Test DateTime", t=>t.Column3),
            };

            List<TestObject> list = new List<TestObject>()
            {
                new TestObject() { Column1 = "Test 1", Column2 = 1, Column3 = new DateTime(2024,06,01) },
                new TestObject() { Column1 = "Test 2", Column2 = 2, Column3 = new DateTime(2023,06,01) },
                new TestObject() { Column1 = "Test 3", Column2 = 3, Column3 = new DateTime(2022,06,01) },
            };


            await application.WriteAsTemplateToAsync(@"UseExcelAsTemplatev2.xlsx", list.ToAsync(), startRow: 1);
        }

        [TestMethod]
        public void ColumnCountTest()
        { 
            for(int i=0; i <700; i++)
            {
                string rowIndex = Cell.ToTextRowIndeX(i);
                if (i == 26)
                    Assert.AreEqual("Z", rowIndex);
                else if (i == 27)
                    Assert.AreEqual("AA", rowIndex);
                else if (i == 52)
                    Assert.AreEqual("AZ", rowIndex);
                else if (i == 53)
                    Assert.AreEqual("BA", rowIndex);
                    

                if(rowIndex.Contains("\\u"))
                    Assert.Fail("Invalid character in row index");

                if(rowIndex.Contains("["))
                    Assert.Fail("Invalid character in row index");
            }
        }


        [TestMethod]
        public void WriteExcel()
        {
            // Assemble
            using (Application application = new Application())
            {

                // Act
                var sheet = application.Workbook.Sheets.Create("Blad 1 sheet");
                
                sheet.Cells[1, 0].Text = "Kolom 1";
                sheet.Cells[2, 0].Text = "Rij 1";
                sheet.Cells[3, 0].Text = "Datum 1";
                sheet.Cells[0, 2].Text = "30-10-2015\nGertjan Trading";

                sheet.Cells[3, 1].Date = DateTime.Now;

                var newSheet = application.Workbook.Sheets.Create("New sheet");
                newSheet.Cells[1, 0].Text = "Sheet 2";


                application.WriteTo(@"WriteExcel.xlsx");
            }
        }

        [TestMethod]
        public async Task WriteExcelTemplate()
        {
            // Assemble
            using Application application = new Application();

            // Act
            var sheet = application.Workbook.Sheets.Create("Blad 1 sheet");
            sheet.ColumnTemplate = new List<ColumnTemplate>()
            {
                new FuncColumnTemplate<TestObject>("Test string", t=>t.Column1),
                new FuncColumnTemplate<TestObject>("Test int", t=>t.Column2),
                new FuncColumnTemplate<TestObject>("Test DateTime", t=>t.Column3),
            };

            List<TestObject> list = new List<TestObject>()
            {
                new TestObject() { Column1 = "Test 1", Column2 = 1, Column3 = new DateTime(2024,06,01) },
                new TestObject() { Column1 = "Test 2", Column2 = 2, Column3 = new DateTime(2023,06,01) },
                new TestObject() { Column1 = "Test 3", Column2 = 3, Column3 = new DateTime(2022,06,01) },
            };

            await sheet.FillFromObjectsAsync(ListToIAsyncEnumerable(list));

            application.WriteTo(@"WriteExcelTemplate.xlsx");

            using Application readApplication = new Application();
            // Act
            readApplication.ReadFile(@"WriteExcelTemplate.xlsx");

            var firstSheet = readApplication.Workbook.Sheets[0];

            // check header
            Assert.AreEqual("Test string", firstSheet.Cells[0, 0].Text);
            Assert.AreEqual("Test int", firstSheet.Cells[0, 1].Text);
            //Assert.AreEqual(new DateTime(2024, 06, 01), firstSheet.Cells[0, 1].Date); For now the read of a datetime format gives a string

            int row = 1;
            foreach (var item in list)
            {
                Assert.AreEqual(item.Column1, firstSheet.Cells[row, 0].Text);
                Assert.AreEqual(item.Column2.ToString(), firstSheet.Cells[row, 1].Text);
                row++;
            }
        }

        [TestMethod]
        public async Task WriteExcelV2Template()
        {
            // Assemble
            using Application application = new Application();

            // Act
            var sheet = application.Workbook.Sheets.Create("Blad 1 sheet");
            sheet.ColumnTemplate = new List<ColumnTemplate>()
            {
                new FuncColumnTemplate<TestObject>("Test string", t=>t.Column1),
                new FuncColumnTemplate<TestObject>("Test int", t=>t.Column2),
                new FuncColumnTemplate<TestObject>("Test DateTime", t=>t.Column3),
            };

            List<TestObject> list = new List<TestObject>()
            {
                new TestObject() { Column1 = "Test 1", Column2 = 1, Column3 = new DateTime(2024,06,01) },
                new TestObject() { Column1 = "Test 2", Column2 = 2, Column3 = new DateTime(2023,06,01) },
                new TestObject() { Column1 = "Test 3", Column2 = 3, Column3 = new DateTime(2022,06,01) },
            };

            await application.FillFromObjectsAndWriteAsync(sheet, ListToIAsyncEnumerable(list), @"WriteExcelV2Template.xlsx", startRow: 1);

            using Application readApplication = new Application();
            // Act
            readApplication.ReadFile(@"WriteExcelV2Template.xlsx");

            var firstSheet = readApplication.Workbook.Sheets[0];

            // check header
            Assert.AreEqual("Test string", firstSheet.Cells[1, 0].Text);
            Assert.AreEqual("Test int", firstSheet.Cells[1, 1].Text);
            //Assert.AreEqual(new DateTime(2024, 06, 01), firstSheet.Cells[0, 1].Date); For now the read of a datetime format gives a string

            int row = 2;
            foreach (var item in list)
            {
                Assert.AreEqual(item.Column1, firstSheet.Cells[row, 0].Text);
                Assert.AreEqual(item.Column2.ToString(), firstSheet.Cells[row, 1].Text);
                row++;
            }
        }

#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
        private async IAsyncEnumerable<T> ListToIAsyncEnumerable<T>(IEnumerable<T> list)
#pragma warning restore CS1998 // Async method lacks 'await' operators and will run synchronously
        {
            foreach (var item in list)
            {
                yield return item;
            }
        }

        class TestObject
        {
            public string Column1 { get; set; }

            public int Column2 { get; set; }
            public DateTime Column3 { get; set; }
        }

        [TestMethod,Ignore]
        public void ReadSharedB5Test()
        {
            // Assemble
            using (Application application = new Application())
            {

                application.ReadFile(@"DataTest/EOD settlement GEM - IceEndex (Pwr-Gas)2016-01-04 190051 (4-1 vs 5-1).xlsx");
            }
                
        }

        [TestMethod]
        public void CellCollectionTest()
        {
            CellCollection cell = new CellCollection(new SharedStrings());
            int row;
            int column;
            cell.ExcelColumnToColumnRowIndex("BA3", out row, out column);

            Assert.AreEqual(2, row);
            Assert.AreEqual(26 * 2, column);
        }

        [TestMethod]
        public void CellCollection2Test()
        {
            CellCollection cell = new CellCollection(new SharedStrings());
            int row;
            int column;
            cell.ExcelColumnToColumnRowIndex("BB3", out row, out column);

            Assert.AreEqual(2, row);
            Assert.AreEqual((26 * 2) + 1, column);
        }

        [TestMethod]
        public void CellCollection3Test()
        {
            CellCollection cell = new CellCollection(new SharedStrings());
            int row;
            int column;
            cell.ExcelColumnToColumnRowIndex("AB3", out row, out column);

            Assert.AreEqual(2, row);
            Assert.AreEqual((26 * 1) + 1, column);
        }

        [TestMethod]
        public void CellCollection4Test()
        {
            CellCollection cell = new CellCollection(new SharedStrings());
            int row;
            int column;
            cell.ExcelColumnToColumnRowIndex("B3", out row, out column);

            Assert.AreEqual(2, row);
            Assert.AreEqual(1, column);
        }

        [TestMethod]
        public void RedTextTest()
        {
            // Assemble
            using Application application = new Application();

            // Act
            application.ReadFile(@"Files/04.Opgenomen-documenten.xlsx");

            // Assert                
            Assert.IsTrue(application.Workbook.Sheets.Count > 3, application.Workbook.Sheets.Count.ToString());
            Sheet sheet1 = application.Workbook.Sheets[0];
            Assert.AreEqual("NEN 1010 Elektrische installaties voor laagspanning", sheet1.Cells[348, 0].Text);
        }
    }
}
