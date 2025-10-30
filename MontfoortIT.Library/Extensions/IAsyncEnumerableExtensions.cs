using System;
using System.Collections.Generic;

namespace MontfoortIT.Library.Extensions
{
    public static class IAsyncEnumerableExtensions
    {
        public static async IAsyncEnumerable<C> Where<C>(this IAsyncEnumerable<C> asyncEnum, Func<C, bool> filter)
        {
            await foreach (var item in asyncEnum)
            {
                if (filter(item))
                    yield return item;
            }
        }


#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
        public static async IAsyncEnumerable<T> ToAsync<T>(this IEnumerable<T> objects)
#pragma warning restore CS1998 // Async method lacks 'await' operators and will run synchronously
        {
            foreach (var obj in objects)
            {
                yield return obj;
            }
        }
    }
}
