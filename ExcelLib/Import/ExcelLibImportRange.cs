using System;
using System.Collections.Generic;
using System.Linq;

using ClosedXML.Excel;

namespace ExcelLib.Import
{
    /// <summary></summary>
    /// <typeparam name="T"></typeparam>
    public sealed class ExcelLibImportRange<T>
    {
        /// <summary></summary>
        private readonly List<ExcelLibImportProperty> _properties;
        /// <summary></summary>
        private readonly Func<ExcelLibImportValuesContainer, ExcelLibResultContainer<T>> _tupleConveter;
        
        /// <summary></summary>
        /// <param name="properties"></param>
        /// <param name="tupleConveter"></param>
        public ExcelLibImportRange(List<ExcelLibImportProperty> properties, Func<ExcelLibImportValuesContainer, ExcelLibResultContainer<T>> tupleConveter)
        {
            if (properties == null) throw new ArgumentNullException("properties");
            if (tupleConveter == null) throw new ArgumentNullException("tupleConveter");

            this._properties = properties;
            this._tupleConveter = tupleConveter;
        }

        /// <summary></summary>
        /// <param name="worksheet"></param>
        /// <param name="firstRow"></param>
        /// <param name="firstColumn"></param>
        /// <returns></returns>
        public List<T> ParseValues(IXLWorksheet worksheet, int firstRow, int firstColumn)
        {
            var propertiesColumnNumbers = this._properties.Select(p => p.RelativeColumnPlace + firstColumn - 1).ToList();

            if (propertiesColumnNumbers.GroupBy(c => c).Any(g => g.Count() > 1)) throw new InvalidOperationException(string.Format("More than one column has unique column number"));

            var list = new List<T>();

            using (var rows = worksheet.RowsUsed(r => propertiesColumnNumbers.Any(c => !r.Cell(c).IsEmpty())))
            {
                if (rows.Any())
                {
                    rows.ForEach(r =>
                    {
                        var result = this._tupleConveter(new ExcelLibImportValuesContainer(this._properties.Select(p => p.ParseValue(worksheet, r.RowNumber(), firstColumn)).ToList()));

                        if (result.IsValid) list.Add(result.Value);
                    });
                }
            }

            return list;
        }
    }
}