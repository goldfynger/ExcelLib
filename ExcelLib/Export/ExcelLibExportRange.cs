using System;
using System.Collections.Generic;
using System.Linq;

using ClosedXML.Excel;

namespace ExcelLib.Export
{
    /// <summary></summary>
    public sealed class ExcelLibExportRange<T>
    {
        /// <summary></summary>
        List<ExcelLibExportValue> _exportValues = new List<ExcelLibExportValue>();

        /// <summary></summary>
        private int _rowsNumber = 0;
        /// <summary></summary>
        private int _columnsNumber = 0;

        /// <summary></summary>
        private Dictionary<int, ExcelLibExportPropertyHeader> _headers;

        /// <summary></summary>
        /// <param name="values"></param>
        /// <param name="convertValuesFunc"></param>
        public ExcelLibExportRange(List<T> values, Func<T, List<ExcelLibExportValue>> convertValuesFunc)
        {
            if (values == null) throw new ArgumentNullException("values");
            if (convertValuesFunc == null) throw new ArgumentNullException("exportValuesFunc");

            var rowCounter = 0;
            var columnNumber = 0;

            Dictionary<int, ExcelLibExportPropertyHeader> headers = null;

            this._exportValues.AddRange(values.Select(v =>
            {
                var internalRowExportValues = convertValuesFunc(v);

                var invalidRowNumbers = internalRowExportValues.GroupBy(vv => vv.GetRelativeColumn()).Where(g => g.Count() > 1).Select(g => g.Key);

                if (invalidRowNumbers.Any()) throw new InvalidOperationException(string.Format("More than one column has column number \"{0}\"", invalidRowNumbers.First()));

                ++rowCounter;

                internalRowExportValues.ForEach(vv => vv.SetRelativeRow(rowCounter));

                if (columnNumber == 0) columnNumber = internalRowExportValues.OrderBy(vv => vv.GetRelativeColumn()).Last().GetRelativeColumn();

                if (headers == null)
                {
                    headers = internalRowExportValues.Select(vv => new { Header = vv.GetValueHeader(), Column = vv.GetRelativeColumn() }).Where(x => x.Header != null).ToDictionary(x => x.Column, x => x.Header);
                }

                return internalRowExportValues;

            }).SelectMany(v => v));

            if (headers != null && headers.Any()) this._headers = headers;

            this._rowsNumber = this._headers == null ? rowCounter : rowCounter + 1;
            this._columnsNumber = columnNumber;
        }

        /// <summary></summary>
        /// <param name="worksheet"></param>
        /// <param name="firstRow"></param>
        /// <param name="firstColumn"></param>
        /// <returns></returns>
        public IXLRange FillValues(IXLWorksheet worksheet, int firstRow, int firstColumn)
        {
            if (firstRow < 1) throw new ArgumentOutOfRangeException("firstRow", string.Format("Invalid row value: {0}", firstRow));
            if (firstColumn < 1) throw new ArgumentOutOfRangeException("firstColumn", string.Format("Invalid column value: {0}", firstColumn));

            var row = this._headers == null || !this._headers.Any() ? firstRow : firstRow + 1;

            if (this._headers != null)
            {
                this._headers.ForEach(p =>
                {
                    var cell = worksheet.Cell(firstRow, firstColumn + p.Key - 1);
                    var header = p.Value;

                    cell.SetValue(ExcelLibService.CheckAndChangeXmlString(header.Header));

                    cell.SetDataType(XLCellValues.Text);
                    
                    var horizontalAlignment = header.HorizontalAlignment;
                    if (horizontalAlignment.HasValue)
                    {
                        cell.Style.Alignment.Horizontal = horizontalAlignment.Value;
                    }

                    var verticalAlignment = header.VerticalAlignment;
                    if (verticalAlignment.HasValue)
                    {
                        cell.Style.Alignment.Vertical = verticalAlignment.Value;
                    }
                });
            }

            this._exportValues.ForEach(v => v.FillValue(worksheet, row, firstColumn));

            return worksheet.Range(firstRow, firstColumn, firstRow + this._rowsNumber - 1, firstColumn + this._columnsNumber - 1);
        }

        /// <summary></summary>
        public int RowsNumber { get { return this._rowsNumber; } }
        /// <summary></summary>
        public int ColumnsNumber { get { return this._columnsNumber; } }
    }
}