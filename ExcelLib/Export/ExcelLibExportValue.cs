using System;

using ClosedXML.Excel;

namespace ExcelLib.Export
{
    /// <summary></summary>
    public abstract class ExcelLibExportValue
    {
        /// <summary></summary>
        /// <param name="relativeRow"></param>
        internal abstract void SetRelativeRow(int relativeRow);

        /// <summary></summary>
        /// <returns></returns>
        internal abstract int GetRelativeRow();
        /// <summary></summary>
        /// <returns></returns>
        internal abstract int GetRelativeColumn();

        /// <summary></summary>
        /// <returns></returns>
        internal abstract ExcelLibExportPropertyHeader GetValueHeader();

        /// <summary></summary>
        /// <param name="worksheet"></param>
        /// <param name="firstRow"></param>
        /// <param name="firstColumn"></param>
        internal abstract void FillValue(IXLWorksheet worksheet, int firstRow, int firstColumn);
    }

    /// <summary></summary>
    public sealed class ExcelLibExportValue<T> : ExcelLibExportValue
    {
        /// <summary></summary>
        private readonly T _value;

        /// <summary></summary>
        private readonly ExcelLibExportProperty<T> _property;

        /// <summary></summary>
        private int _relativeRow;

        /// <summary></summary>
        /// <param name="value"></param>
        /// <param name="property"></param>
        /// <param name="row"></param>
        public ExcelLibExportValue(T value, ExcelLibExportProperty<T> property)
        {
            if (value == null) throw new ArgumentNullException("value");
            if (property == null) throw new ArgumentNullException("property");

            this._value = value;
            this._property = property;
        }

        /// <summary></summary>
        /// <param name="relativeRow"></param>
        internal override void SetRelativeRow(int relativeRow)
        {
            if (relativeRow < 1) throw new ArgumentOutOfRangeException("relativeRow", string.Format("Invalid row value: {0}", relativeRow));

            this._relativeRow = relativeRow;
        }

        /// <summary></summary>
        /// <returns></returns>
        internal override int GetRelativeRow()
        {
            return this._relativeRow;
        }
        /// <summary></summary>
        /// <returns></returns>
        internal override int GetRelativeColumn()
        {
            return this._property.RelativeColumnPlace;
        }

        /// <summary></summary>
        /// <returns></returns>
        internal override ExcelLibExportPropertyHeader GetValueHeader()
        {
            return this._property.Header;
        }

        /// <summary></summary>
        /// <param name="worksheet"></param>
        /// <param name="firstRow"></param>
        /// <param name="firstColumn"></param>
        internal override void FillValue(IXLWorksheet worksheet, int firstRow, int firstColumn)
        {
            var row = firstRow + this._relativeRow - 1;
            var column = firstColumn + this._property.RelativeColumnPlace - 1;

            var cell = worksheet.Cell(row, column);

            cell.SetValue(ExcelLibService.CheckAndChangeXmlString(this._property.ValueConverter == null ? this._value.ToString() : this._property.ValueConverter(this._value)));

            var dataType = this._property.GetDataType(this._value);
            if (dataType.HasValue)
            {
                cell.SetDataType(dataType.Value);
            }

            var horizontalAlignment = this._property.GetHorizontalAlignment(this._value);
            if (horizontalAlignment.HasValue)
            {
                cell.Style.Alignment.Horizontal = horizontalAlignment.Value;
            }

            var verticalAlignment = this._property.GetVerticalAlignment(this._value);
            if (verticalAlignment.HasValue)
            {
                cell.Style.Alignment.Vertical = verticalAlignment.Value;
            }

            var color = this._property.GetBackgroundColor(this._value);
            if (color.HasValue)
            {
                cell.Style.Fill.BackgroundColor = XLColor.FromColor(color.Value);
            }
        }
    }
}