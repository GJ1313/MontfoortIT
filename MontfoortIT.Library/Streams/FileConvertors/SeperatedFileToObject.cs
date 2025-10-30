using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace MontfoortIT.Library.Streams.FileConvertors
{
    public abstract class SeperatedFileToObject<T> : SeperatedFileToBase
    {

        public List<Templates.FuncColumnTemplate<T>> Template { get; set; }


        public async Task<IEnumerable<T>> ConvertAsync(TextReader reader, Func<T> constructor)
        {
            var objectWriter = new ObjectWriter<T>(Template, constructor);
            await ConvertAsync(reader, objectWriter);

            return objectWriter.Objects;
        }

        public async Task<IEnumerable<T>> ConvertAsync(Stream file, Func<T> constructor)
        {
            var objectWriter = new ObjectWriter<T>(Template, constructor);
            await ConvertAsync(file, objectWriter);

            return objectWriter.Objects;
        }

    }
}
