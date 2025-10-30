using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace MontfoortIT.Office.Excel.Word
{
    public class Document
    {
        private Package _package;
        private IMergeFieldResolver _resolver;
        private string _sjabloonPath;
        private Stream _sjabloonStream;
        private static string _storePath;
        private static Stream _storeStream;

        private Document(string filePath)
        {
            _sjabloonPath = filePath;
        }

        public Document(Stream sjabloonStream)
        {
            _sjabloonStream = sjabloonStream;
        }

        public static Document Load(string filePath)
        {
            return new Document(filePath);
        }

        public static Document Load(Stream sjabloonStream)
        {
            return new Document(sjabloonStream);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="resolver"></param>
        /// <returns></returns>
        /// <remarks>Dispose the stream when finished</remarks>
        public Stream Merge(IMergeFieldResolver resolver)
        {
            _resolver = resolver;

            _storeStream = new MemoryStream();

            Merge();
            _storeStream.Position = 0;
            return _storeStream;
        }

        public void Merge(IMergeFieldResolver resolver, string storePath)
        {
            _resolver = resolver;
            _storePath = storePath;

            Merge();
        }

        private void Merge()
        {
            using (Package package = OpenPackage())
            {
                _package = package;

                using (Package savePackage = GetNewPackage())
                {
                    foreach (var mergePart in _package.GetParts())
                    {
                        string contentTypeDest = ConvertContentType(mergePart.ContentType);

                        PackagePart newPart = savePackage.CreatePart(mergePart.Uri, contentTypeDest, mergePart.CompressionOption);
                        if (mergePart.ContentType == ContentTypes.WordXmlTemplatePart)
                        {
                            XDocument resultDoc = ProcessPart(mergePart);
                            using (Stream newStream = newPart.GetStream(FileMode.Create))
                            using (XmlWriter newWriter = XmlWriter.Create(newStream))
                            {
                                resultDoc.WriteTo(newWriter);
                            }
                        }
                        else
                        {

                            using (Stream newStream = newPart.GetStream(FileMode.Create))
                            using (Stream originalStream = mergePart.GetStream())
                            {
                                originalStream.CopyTo(newStream);
                            }
                        }
                    }

                    AddRelations(savePackage);

                    savePackage.Flush();
                }
                _package = null;
            }
        }

        private string ConvertContentType(string originalContentType)
        {
            switch (originalContentType)
            {
                case "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml":
                    return "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

            }
            return originalContentType;
        }

        private static Package GetNewPackage()
        {
            if (string.IsNullOrEmpty(_storePath))
            {
                return Package.Open(_storeStream, FileMode.Create);
            }
            else
            {
                return Package.Open(_storePath, FileMode.Create);
            }
        }

        private Package OpenPackage()
        {
            if(string.IsNullOrEmpty(_sjabloonPath))
                return Package.Open(_sjabloonStream, FileMode.Open, FileAccess.Read);
            else
                return Package.Open(_sjabloonPath, FileMode.Open, FileAccess.Read);
        }

        private XDocument ProcessPart(PackagePart part)
        {
            XDocument partDoc;
            
            using(Stream partStream = part.GetStream())
            using(XmlReader partReader = XmlReader.Create(partStream))
            {
                partDoc = XDocument.Load(partReader);
            }

            //<w:fldChar w:fldCharType="begin"/>
            List<XElement> mergeFields = partDoc.Descendants(Namespaces.WordProcessingXml + "fldChar")
                .Where(e=>e.Attribute(Namespaces.WordProcessingXml + "fldCharType").Value=="begin")
                .Select(e=>e.Parent.Parent).ToList();
            foreach (var replaceFields in mergeFields)
            {
                ProcessField(replaceFields);                                
            }
            mergeFields = partDoc.Descendants(Namespaces.WordProcessingXml + "fldSimple").Select(e => e.Parent).ToList();
            foreach (var replaceFields in mergeFields)
            {
                ProcessFieldSimple(replaceFields);
            }

            return partDoc;

            //if(mergeFields.Any())
            //{
            //    using(Stream stream = part.GetStream(FileMode.Open))
            //    using(XmlWriter partWriter = XmlWriter.Create(stream))
            //    {
            //        partDoc.WriteTo(partWriter);
            //    }
            //}
            
            //return part;
        }

        private void ProcessFieldSimple(XElement replaceFields)
        {
            string fieldMerge = replaceFields.Element(Namespaces.WordProcessingXml + "fldSimple")
                .Attribute(Namespaces.WordProcessingXml + "instr").Value.Substring(11).Trim();

            if (fieldMerge.EndsWith("MERGEFORMAT"))
            {
                fieldMerge = fieldMerge.Substring(0, fieldMerge.Length - 14).Trim();
            }

            ProcessField(replaceFields, fieldMerge);
        }

        private void ProcessField(XElement replaceFields)
        {
            string field = null;
            string fieldText;
            foreach (var alina in replaceFields.Elements().ToList())
            {
                XElement firstElement = (XElement)alina.FirstNode;

                                
            }


            foreach (var instText in replaceFields.Descendants(Namespaces.WordProcessingXml + "instrText").ToList())
            {
                if (instText.Value.StartsWith("MERGEFIELD") && instText.Value.Length > 12)
                {
                    field = instText.Value.Substring(11).Trim();
                    fieldText = _resolver.GetFieldValue(field);
                    instText.ReplaceWith(new XElement(Namespaces.WordProcessingXml + "t", fieldText));
                }
            }
        }

        private void ProcessField(XElement replaceFields, string field)
        {

            if (field == null)
                return;
            string fieldText = _resolver.GetFieldValue(field);

            replaceFields.RemoveNodes();
            replaceFields.Add(new XElement(Namespaces.WordProcessingXml + "r", new XElement(Namespaces.WordProcessingXml + "t", fieldText)));
        }

        public void SaveAs(string filePath)
        {
            using(FileStream saveStream = File.Open(filePath, FileMode.Create))
            {
                Package savePackage = Package.Open(saveStream, FileMode.Create, FileAccess.Write);
                foreach (var mergePart in _package.GetParts())
                {
                    PackagePart newPart = savePackage.CreatePart(mergePart.Uri, mergePart.ContentType, mergePart.CompressionOption);
                    using (Stream newStream = newPart.GetStream(FileMode.Create))
                    using (Stream originalStream = mergePart.GetStream())
                    {
                        originalStream.CopyTo(newStream);
                    }
                }

                AddRelations(savePackage);

                savePackage.Flush();
            }
        }

        private void AddRelations(Package savePackage)
        {
            //foreach (var old in savePackage.GetRelationships().ToList())
            //{
            //    savePackage.DeleteRelationship(old.Id);
            //}

            //foreach (var relation in _package.GetRelationships())
            //{
            //    if (savePackage.GetRelationship(relation.Id) != null)
            //        savePackage.CreateRelationship(relation.TargetUri, relation.TargetMode, relation.RelationshipType, relation.Id);
            //}
        }
    }
}
