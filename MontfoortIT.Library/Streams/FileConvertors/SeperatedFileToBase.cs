using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MontfoortIT.Library.Streams.FileConvertors
{
    public abstract class SeperatedFileToBase
    {
        protected abstract char Seperator { get; }
        protected virtual char[] ExcludeChars { get { return new char[0]; } }

        protected virtual bool ContinueRead
        {
            get { return false; }
        }

        public async Task ConvertAsync(TextReader streamReader, IStructureWriter to, Encoding fromEncoding = null)
        {
            to.WriteStartDocument();
            to.WriteStartTable();

            List<int> excludeChars = ExcludeChars.Select(excludeChar => (int)excludeChar).ToList();

            //int c = streamReader.Read();
            string line = "start";
            while (line!=null)
            {
                to.WriteStartRow();

                line = await streamReader.ReadLineAsync();
                ConvertLineToColumns(to, excludeChars, line);

                to.WriteEndRow();
            }

            to.WriteEndTable();
            to.WriteEndDocument();
        }

        public string[] ConvertLineToColumns(string line)
        {
            List<int> excludeChars = ExcludeChars.Select(excludeChar => (int)excludeChar).ToList();
            var inMemWriter = new InMemoryStructureWriter();
            ConvertLineToColumns(inMemWriter, excludeChars, line);

            return inMemWriter.GetColumns();    
        }

        private void ConvertLineToColumns(IStructureWriter to, List<int> excludeChars, string line)
        {
            if (!string.IsNullOrEmpty(line))
            {
                int cCount = 0;
                char? c = line[cCount];
                while (c.HasValue && c >= 0 && c != '\n')
                {
                    to.WriteStartColumn();

                    StringBuilder columnBuilder = new StringBuilder();
                    //StringBuilder strBuilder = new StringBuilder();
                    while (c.HasValue && (c != Seperator || ContinueRead) && c != '\n')
                    {
                        if (excludeChars.Contains(c.Value))
                            c = GetNextChar(line, ref cCount);

                        if (c.HasValue)
                        {
                            bool append = ProcessChar(c.Value);

                            if (append && c != '\r' && c != '\n')
                                columnBuilder.Append(c);
                        }

                        // strBuilder.Append(ch);

                        c = GetNextChar(line, ref cCount);
                    }

                    string text = columnBuilder.ToString();
                    text = CleanText(text);
                    to.WriteString(text);

                    to.WriteEndColumn();

                    if (c == Seperator)
                        c = GetNextChar(line, ref cCount);
                }
            }
        }

        public Task ConvertAsync(Stream from, IStructureWriter to, Encoding fromEncoding = null)
        {
            Encoding utf = Encoding.UTF8;
            if (fromEncoding == null)
                fromEncoding = utf;

            StreamReader streamReader = new StreamReader(from, fromEncoding);

            //Encoding utf = Encoding.GetEncoding("ISO-8859-15");
            // .NET core does not support the ISO-8859-15.... https://stackoverflow.com/questions/398621/system-text-encoding-getencodingiso-8859-1-throws-platformnotsupportedexcept
            return ConvertAsync(streamReader, to, fromEncoding);

        }

        private static char? GetNextChar(string line, ref int cCount)
        {            
            cCount++;
            if(cCount<line.Length)
                return line[cCount];
            return null;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ch"></param>
        /// <returns>If the char should be added to the column</returns>
        protected virtual bool ProcessChar(char ch)
        {
            return true;
        }


        private string CleanText(string text)
        {
            return new string(text.Where(c => c > 8 && c < 11 || c > 12 && c < 14 || c > 26).ToArray());
        }

    }
}