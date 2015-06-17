using System;
using System.Drawing;

using ClosedXML.Excel;

namespace ExcelLib.Export
{
    /// <summary></summary>
    public sealed class ExcelLibExportString
    {
        /// <summary></summary>
        private readonly string _value;
        /// <summary></summary>
        private readonly int _cellsCount;
        /// <summary></summary>
        private readonly Color? _backgroundColor;
        /// <summary></summary>
        private readonly XLAlignmentHorizontalValues? _horizontalAlignment;
        /// <summary></summary>
        private readonly XLAlignmentVerticalValues? _verticalAlignment;

        /// <summary></summary>
        /// <param name="value"></param>
        /// <param name="cellsCount"></param>
        /// <param name="backgroundColor"></param>
        /// <param name="horizontalAlignment"></param>
        /// <param name="verticalAlignment"></param>
        public ExcelLibExportString(string value, int cellsCount, Color? backgroundColor = null, XLAlignmentHorizontalValues? horizontalAlignment = null, XLAlignmentVerticalValues? verticalAlignment = null)
        {
            if (value == null) throw new ArgumentNullException("value");

            if (cellsCount < 1) throw new ArgumentOutOfRangeException("cellsCount");

            this._value = value;
            this._cellsCount = cellsCount;
            this._backgroundColor = backgroundColor;
            this._horizontalAlignment = horizontalAlignment;
            this._verticalAlignment = verticalAlignment;
        }

        /// <summary></summary>
        /// <param name="worksheet"></param>
        /// <param name="row"></param>
        /// <param name="firstColumn"></param>
        public void FillString(IXLWorksheet worksheet, int row, int firstColumn)
        {
            var cell = worksheet.Cell(row, firstColumn);

            cell.SetValue(ExcelLibService.CheckAndChangeXmlString(this._value));

            cell.SetDataType(XLCellValues.Text);

            if (this._horizontalAlignment.HasValue)
            {
                cell.Style.Alignment.Horizontal = this._horizontalAlignment.Value;
            }

            if (this._verticalAlignment.HasValue)
            {
                cell.Style.Alignment.Vertical = this._verticalAlignment.Value;
            }

            if (this._backgroundColor.HasValue)
            {
                cell.Style.Fill.BackgroundColor = XLColor.FromColor(this._backgroundColor.Value);
            }

            if (this._cellsCount > 1)
            {
                worksheet.Range(row, firstColumn, row, firstColumn + this._cellsCount - 1).Merge();
            }
        }
    }
}