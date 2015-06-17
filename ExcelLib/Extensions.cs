using System;
using System.Collections.Generic;

namespace ExcelLib
{
    /// <summary></summary>
    internal static class Extensions
    {
        /// <summary></summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source"></param>
        /// <param name="action"></param>
        public static void ForEach<T>(this IEnumerable<T> source, Action<T> action)
        {
            foreach (T element in source)
            {
                action(element);
            }
        }
    }
}