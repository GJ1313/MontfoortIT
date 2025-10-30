using MontfoortIT.Office.Excel.Csv;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace MontfoortIT.Office.Excel
{
    public interface ISheet
    {
        List<ColumnTemplate> ColumnTemplate { get; set; }

        Task WriteFromObjectsAsync<D>(IAsyncEnumerable<D> objects, Stream fileStream, Options options, Encoding encoding, bool writeHeaderRow = true);
        Task WriteFromObjectsAsync<D>(IEnumerable<D> objects, Stream fileStream, Options options, Encoding encoding, bool writeHeaderRow = true);
    }
}