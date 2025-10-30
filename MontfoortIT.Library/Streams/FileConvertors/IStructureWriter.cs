namespace MontfoortIT.Library.Streams.FileConvertors
{
    public interface IStructureWriter
    {
        void WriteStartDocument();
        void WriteStartTable();
        void WriteStartRow();
        void WriteStartColumn();
        void WriteString(string text);
        void WriteEndColumn();
        void WriteEndDocument();
        void WriteEndTable();
        void WriteEndRow();
    }
}