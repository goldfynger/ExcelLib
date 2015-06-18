using System;
using System.Collections.Generic;

namespace ExcelLib.Import
{
    public sealed class ExcelLibImportValuesContainer
    {
        /// <summary></summary>
        private Dictionary<Guid, ExcelLibImportValue> _dictionary = new Dictionary<Guid, ExcelLibImportValue>();

        /// <summary></summary>
        /// <param name="values"></param>
        internal ExcelLibImportValuesContainer(List<ExcelLibImportValue> values)
        {
            values.ForEach(v => this.AddValue(v));
        }

        /// <summary></summary>
        /// <param name="value"></param>
        internal void AddValue(ExcelLibImportValue value)
        {
            this._dictionary.Add(value.Guid, value);
        }

        /// <summary></summary>
        /// <param name="guid"></param>
        /// <returns></returns>
        internal ExcelLibImportValue GetValue(Guid guid)
        {
            return this._dictionary[guid];
        }
    }
}