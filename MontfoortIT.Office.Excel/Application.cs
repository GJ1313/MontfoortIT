using MontfoortIT.Library.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Xml;
using System.Xml.Linq;

namespace MontfoortIT.Office.Excel
{
    public class Application:IDisposable
    {
        private Workbook _workbook;

        private SharedStrings _sharedStrings;
        private Styles _styles;
        private Package _originalPackage;

        internal readonly Func<NumberFormat,bool, string> ConvertNumberFormatToSAttribute;

        public string ContentType
        {
            get
            {
                return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            }
        }

        internal SharedStrings SharedStrings
        {
            get
            {
                if (_sharedStrings == null)
                {
                    _sharedStrings = new SharedStrings();
                }
                return _sharedStrings;
            }
        }

        internal Styles Styles
        {
            get
            {
                if (_styles == null)
                    _styles = new Styles();
                return _styles;
            }
        }


        public Workbook Workbook
        {
            get
            {
                if (_workbook == null)
                {
                    _workbook = new Workbook(this);
                }
                return _workbook;
            }
        }

        public Application()
        {
            ConvertNumberFormatToSAttribute = (n,header) => ((int)n).ToString();
        }

        public Application(Func<NumberFormat, bool, string> convertNumberFormatToSAttribute)
        {
            ConvertNumberFormatToSAttribute = convertNumberFormatToSAttribute;
        }

        public Task WriteAsTemplateToAsync(string path, int startRow = 0)
        {
            Package newPackage = CreateExcelWorkbookPackage(path);

            return WriteAsTemplateToAsync<object>(newPackage, null, startRow);
        }

        public Task WriteAsTemplateToAsync(Stream stream, int startRow = 0)
        {
            Package newPackage = CreateExcelWorkbookPackage(stream);

            return WriteAsTemplateToAsync<object>(newPackage,null,startRow);
        }

        public Task WriteAsTemplateToAsync<T>(string path, IAsyncEnumerable<T> objects, int startRow = 0)
        {
            Package newPackage = CreateExcelWorkbookPackage(path);

            return WriteAsTemplateToAsync(newPackage, objects, startRow);
        }

        public Task WriteAsTemplateToAsync<T>(Stream stream, IAsyncEnumerable<T> objects, int startRow = 0)
        {
            Package newPackage = CreateExcelWorkbookPackage(stream);

            return WriteAsTemplateToAsync(newPackage, objects, startRow);
        }

        private async Task WriteAsTemplateToAsync<T>(Package newPackage, IAsyncEnumerable<T>? objects, int startRow=0)
        {
            if (_originalPackage == null)
                throw new NotSupportedException("Template not set, use ReadFile to load a template");

            Uri workSheetUri = PackUriHelper.CreatePartUri(new
                                      Uri("xl/worksheets/sheet1.xml", UriKind.Relative));

            Uri sharedStringsUri = PackUriHelper.CreatePartUri(new
                  Uri("xl/sharedStrings.xml", UriKind.Relative));
            
            List<XDocument> xSheets = new List<XDocument>();
            var extraSheets = Workbook.Sheets.Skip(1).ToList();
            foreach (var sheet in extraSheets)
            {
                xSheets.Add(sheet.CreateDoc());
            }

            PackagePart? sharedStringPart = null;
            foreach (var part in _originalPackage.GetParts())
            {
                if (newPackage.PartExists(part.Uri) || part.Uri.OriginalString == "/_rels/.rels" || part.Uri.OriginalString == "/xl/_rels/workbook.xml.rels")
                    continue;

                PackagePart newPart = newPackage.CreatePart(part.Uri, part.ContentType, CompressionOption.Normal);

                if (part.Uri == sharedStringsUri)
                {
                    sharedStringPart = newPart; // needs to be written at end
                    continue;
                }

                using Stream partStream = newPart.GetStream();
                
                if (part.Uri == workSheetUri)
                {
                    var appSheet = Workbook.Sheets[0];
                    XDocument xSheet;
                    if (objects==null)
                        xSheet = appSheet.WriteInTemplate();
                    else
                        xSheet = await appSheet.FillObjectsInTemplateAsync(objects,startRow);

                    using XmlWriter xmlWriter = XmlWriter.Create(partStream);
                    xSheet.Save(xmlWriter);
                }                
                else if(part.Uri.OriginalString == "/xl/workbook.xml")
                {
                    if(xSheets.Count > 0)
                    {
                        XDocument workbook;
                        using (Stream oldWorkBook = part.GetStream())
                        {
                            workbook = XDocument.Load(oldWorkBook);
                        }

                        XElement sheetsElement = Workbook.GetSheetsXml(workbook);
                        foreach (var extraSheet in extraSheets)
                        {
                            Workbook.AddSheetToXml(sheetsElement, extraSheet);
                        }
                                                    
                        using (XmlWriter xmlWriter = XmlWriter.Create(partStream))
                        {
                            workbook.Save(xmlWriter);
                        }
                    }
                    else
                    {
                        Stream oldStream = part.GetStream();
                        await oldStream.CopyToAsync(partStream);
                    }
                }
                else
                {

                    Stream oldStream = part.GetStream();
                    await oldStream.CopyToAsync(partStream);
                }

                if (!part.ContentType.Contains("application/vnd.openxmlformats-package.relationships+xml"))
                {
                    foreach (var rel in part.GetRelationships())
                    {
                        newPart.CreateRelationship(rel.TargetUri, rel.TargetMode, rel.RelationshipType, rel.Id);
                    }
                }
            }

            if(sharedStringPart != null)
            {
                using Stream sharedStreamWriter = sharedStringPart.GetStream();
                XDocument xSharedStrings = SharedStrings.CreateDoc();
                xSharedStrings.Save(sharedStreamWriter);                
            }

            foreach (var packRel in _originalPackage.GetRelationships())
            {
                newPackage.CreateRelationship(packRel.TargetUri, packRel.TargetMode, packRel.RelationshipType, packRel.Id);
            }

            foreach (var newSheet in xSheets)
            {
                AddExcelPart(newPackage, "worksheet", newSheet);
            }

            newPackage.Flush();
            newPackage.Close();
        }

        public void WriteTo(Stream streamToWriteIn)
        {
            //Create the Excel package.
            Package xlPackage = CreateExcelWorkbookPackage(streamToWriteIn);

            //Create the XML documents.
            WriteTo(xlPackage);
        }

        public Stream GetExcelStream()
        {
            MemoryStream memStream = new MemoryStream();
            WriteTo(memStream);
            memStream.Position = 0;
            return memStream;
        }

        public void WriteTo(string path)
        {
            //Create the Excel package.
            Package xlPackage = CreateExcelWorkbookPackage(path);

            //Create the XML documents.
            WriteTo(xlPackage);
        }

        public async Task FillFromObjectsAndWriteAsync<D>(Sheet sheetToFill, IEnumerable<D> objects, Stream stream, int startRow = 0, bool writeHeaderRow = true)
        {
            //Create the Excel package.
            Package xlPackage = CreateExcelWorkbookPackage(stream);
            await FillFromObjectsAndWriteInternAsync(sheetToFill, objects.ToAsync(), startRow, writeHeaderRow, xlPackage);
        }

        public async Task FillFromObjectsAndWriteAsync<D>(Sheet sheetToFill, IAsyncEnumerable<D> objects, Stream stream, int startRow = 0, bool writeHeaderRow = true)
        {
            //Create the Excel package.
            Package xlPackage = CreateExcelWorkbookPackage(stream);

            await FillFromObjectsAndWriteInternAsync(sheetToFill, objects, startRow, writeHeaderRow, xlPackage);
        }

        public async Task FillFromObjectsAndWriteAsync<D>(Sheet sheetToFill, IEnumerable<D> objects, string path, int startRow = 0, bool writeHeaderRow = true)
        {
            //Create the Excel package.
            Package xlPackage = CreateExcelWorkbookPackage(path);
            await FillFromObjectsAndWriteInternAsync(sheetToFill, objects.ToAsync(), startRow, writeHeaderRow, xlPackage);
        }

        public async Task FillFromObjectsAndWriteAsync<D>(Sheet sheetToFill, IAsyncEnumerable<D> objects, string path, int startRow = 0, bool writeHeaderRow = true)
        {
            //Create the Excel package.
            Package xlPackage = CreateExcelWorkbookPackage(path);

            await FillFromObjectsAndWriteInternAsync(sheetToFill, objects, startRow, writeHeaderRow, xlPackage);
        }

        private async Task FillFromObjectsAndWriteInternAsync<D>(Sheet sheetToFill, IAsyncEnumerable<D> objects, int startRow, bool writeHeaderRow, Package xlPackage)
        {
            XDocument workbookDoc = Workbook.CreateDoc();

            List<XDocument> sheets = new List<XDocument>();
            foreach (var sheet in Workbook.Sheets)
            {
                if (sheet == sheetToFill)
                {
                    var sheetDoc = await sheet.CreateDocAndFillAsync(objects, startRow, writeHeaderRow);
                    sheets.Add(sheetDoc);
                }
                else
                    sheets.Add(sheet.CreateDoc());
            }

            WriteOtherComponents(xlPackage, workbookDoc, sheets);
        }

        public void WriteTo(Package xlPackage)
        {
            XDocument workbookDoc = Workbook.CreateDoc();
            //CreateExcelXML("workbook");

            List<XDocument> sheets = new List<XDocument>();
            foreach (var sheet in Workbook.Sheets)
            {
                sheets.Add(sheet.CreateDoc());
            }

            WriteOtherComponents(xlPackage, workbookDoc, sheets);
        }

        private void WriteOtherComponents(Package xlPackage, XDocument workbookDoc, List<XDocument> sheets)
        {
            XDocument sharedstringsDoc = SharedStrings.CreateDoc();
            XDocument stylesDoc = Styles.CreateDoc();

            //Add the parts to the Excel package.
            if (xlPackage != null)
            {
                //Add the workbook part.
                AddExcelPart(xlPackage, "workbook", workbookDoc);

                //Add the worksheet part.
                foreach (var sheet in sheets)
                {
                    AddExcelPart(xlPackage, "worksheet", sheet);
                }

                //Add the sharedstrings part.
                AddExcelPart(xlPackage, "sharedstrings", sharedstringsDoc);

                AddExcelPart(xlPackage, "styles", stylesDoc);
            }

            //Save the changes, and then close the package.
            if (xlPackage != null)
            {
                xlPackage.Flush();
                xlPackage.Close();
            }
        }

        private static Package CreateExcelWorkbookPackage(string path)
        {
            //Create a new Excel workbook package on the
            //desktop of the user by using the specified name.)
            return Package.Open(path, FileMode.Create, FileAccess.ReadWrite);
        }

        private static Package CreateExcelWorkbookPackage(Stream stream)
        {
            //Create a new Excel workbook package on the
            //desktop of the user by using the specified name.)
            return Package.Open(stream, FileMode.Create, FileAccess.ReadWrite);
        }

        private void AddExcelPart(Package fPackage, string part,
XDocument xDoc)
        {
            switch (part)
            {
                case "workbook":

                    Uri workBookUri = PackUriHelper.CreatePartUri(new
                            Uri("xl/workbook.xml", UriKind.Relative));

                    //Create the workbook part.
                    PackagePart wbPart =
                            fPackage.CreatePart(workBookUri, Namespaces.NsWorkbook, CompressionOption.Normal);

                    //Write the workbook XML to the workbook part.
                    Stream workbookStream =
                            wbPart.GetStream(FileMode.Create, FileAccess.Write);

                    using (XmlWriter xmlWriter = XmlWriter.Create(workbookStream))
                    {
                        xDoc.Save(xmlWriter);
                    }

                    //Create the relationship for the workbook part.
                    fPackage.CreateRelationship(workBookUri, TargetMode.Internal, Namespaces.WorkbookRelationshipType, "rId1");

                    break;

                case "worksheet":

                    int sheetIndex = 1;
                    Uri workSheetUri = PackUriHelper.CreatePartUri(new
                          Uri("xl/worksheets/sheet1.xml", UriKind.Relative));
                    while (fPackage.PartExists(workSheetUri))
                    {
                        sheetIndex++;
                        workSheetUri = PackUriHelper.CreatePartUri(new
                          Uri(string.Format("xl/worksheets/sheet{0}.xml", sheetIndex), UriKind.Relative));
                    }


                    //Create the workbook part.
                    PackagePart wsPart =
                            fPackage.CreatePart(workSheetUri, Namespaces.NsWorksheet, CompressionOption.Normal);

                    //Write the workbook XML to the workbook part.
                    Stream worksheetStream =
                              wsPart.GetStream(FileMode.Create,
                    FileAccess.Write);

                    using (XmlWriter xmlWriter = XmlWriter.Create(worksheetStream))
                    {
                        xDoc.Save(xmlWriter);
                    }

                    //Create the relationship for the workbook part.
                    Uri wsworkbookPartUri = PackUriHelper.CreatePartUri(new
                            Uri("xl/workbook.xml", UriKind.Relative));
                    PackagePart wsworkbookPart = fPackage.GetPart(wsworkbookPartUri);

                    wsworkbookPart.CreateRelationship(workSheetUri,
                              TargetMode.Internal, Namespaces.WorksheetRelationshipType, GenerateNewRelationId(wsworkbookPart, sheetIndex));

                    break;

                case "sharedstrings":
                    string nsSharedStrings = ContentTypes.XlSharedStrings;
                    string sharedStringsRelationshipType =
"http://schemas.openxmlformats.org" + "/officeDocument/2006/relationships/sharedStrings";

                    Uri sharedStringsUri = PackUriHelper.CreatePartUri(new
                          Uri("xl/sharedStrings.xml", UriKind.Relative));

                    //Create the workbook part.
                    PackagePart sharedStringsPart = fPackage.CreatePart(sharedStringsUri,
nsSharedStrings, CompressionOption.Normal);

                    //Write the workbook XML to the workbook part.
                    Stream sharedStringsStream =
                            sharedStringsPart.GetStream(FileMode.Create,
FileAccess.Write);
                    using (XmlWriter xmlWriter = XmlWriter.Create(sharedStringsStream))
                    {
                        xDoc.Save(xmlWriter);
                    }

                    //Create the relationship for the workbook part.
                    Uri ssworkbookPartUri = PackUriHelper.CreatePartUri(new
                            Uri("xl/workbook.xml", UriKind.Relative));
                    PackagePart ssworkbookPart =
                    fPackage.GetPart(ssworkbookPartUri);
                    ssworkbookPart.CreateRelationship(sharedStringsUri,
                            TargetMode.Internal,
                    sharedStringsRelationshipType, "rShared1");

                    break;

                case "styles":
                    Uri stylesUri = PackUriHelper.CreatePartUri(new
                          Uri("xl/styles.xml", UriKind.Relative));

                    //Create the workbook part.
                    PackagePart stylesPart = fPackage.CreatePart(stylesUri, Namespaces.StylesType, CompressionOption.Normal);

                    //Write the workbook XML to the workbook part.
                    Stream stylesStream = stylesPart.GetStream(FileMode.Create, FileAccess.Write);

                    using (XmlWriter xmlWriter = XmlWriter.Create(stylesStream))
                    {
                        xDoc.Save(xmlWriter);
                    }

                    //Create the relationship for the workbook part.
                    Uri stylesWorkbookPartUri = PackUriHelper.CreatePartUri(new
                            Uri("xl/workbook.xml", UriKind.Relative));
                    PackagePart stylesWorkbookPart = fPackage.GetPart(stylesWorkbookPartUri);

                    stylesWorkbookPart.CreateRelationship(stylesUri, TargetMode.Internal, Namespaces.StylesRelationshipNS, "rStyles1");
                    break;
            }
        }

        private string GenerateNewRelationId(PackagePart fPackage, int sheetIndex)
        {
            int number = sheetIndex;
            string relationId = "rIdNew" + number;

            while(fPackage.GetRelationships().Any(r=>r.Id == relationId))
            {
                number++;
                relationId = "rIdNew" + number;
            }

            return relationId;
        }
        

        public void ReadFile(Stream fileStream, int emptyRowsToIgnore = -1)
        {
            Package package = Package.Open(fileStream);
            ReadPackage(package, emptyRowsToIgnore);
        }

        public void ReadFile(string filePath)
        {
            Package package = Package.Open(filePath);            
            ReadPackage(package);
        }

        private void ReadPackage(Package package, int emptyRowsToIgnore = -1)
        {
            _originalPackage = package;
            PackagePartCollection parts = package.GetParts();

            _sharedStrings = new SharedStrings();
            _sharedStrings.Read(parts.Where(p => p.ContentType == ContentTypes.XlSharedStrings).SingleOrDefault());

            IEnumerable<PackagePart> sheets = parts.Where(p => p.ContentType == ContentTypes.XlSheet);
            Workbook.Sheets.Read(sheets);
        }

        public void Dispose()
        {
            if (_originalPackage != null)
            {
                _originalPackage.Close();
                ((IDisposable)_originalPackage).Dispose();
            }

            foreach (var sheet in Workbook.Sheets)
            {
                sheet.Dispose();
            }
        }

    }
}

