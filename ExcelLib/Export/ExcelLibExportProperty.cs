using System;
using System.Drawing;

using ClosedXML.Excel;

namespace ExcelLib.Export
{
    /// <summary></summary>
    /// <typeparam name="T"></typeparam>
    public sealed class ExcelLibExportProperty<T>
    {
        /// <summary></summary>
        private readonly int _relativeColumnPlace;
        /// <summary></summary>
        private readonly Func<T, string> _valueConverter;
        /// <summary></summary>
        private readonly ExcelLibExportPropertyHeader _header;
        /// <summary></summary>
        private readonly ExcelLibExportPropertyParameter<T, Color?> _backgroundColor;
        /// <summary></summary>
        private readonly ExcelLibExportPropertyParameter<T, XLAlignmentHorizontalValues?> _horizontalAlignment;
        /// <summary></summary>
        private readonly ExcelLibExportPropertyParameter<T, XLAlignmentVerticalValues?> _verticalAlignment;
        /// <summary></summary>
        private readonly ExcelLibExportPropertyParameter<T, XLCellValues?> _dataType;

        /// <summary></summary>
        /// <param name="relativeColumnPlace"></param>
        /// <param name="valueConverter"></param>
        /// <param name="header"></param>
        /// <param name="backgroundColor"></param>
        /// <param name="horizontalAlignment"></param>
        /// <param name="verticalAlignment"></param>
        /// <param name="dataType"></param>
        public ExcelLibExportProperty(int relativeColumnPlace, Func<T, string> valueConverter = null, ExcelLibExportPropertyHeader header = null,
            ExcelLibExportPropertyParameter<T, Color?> backgroundColor = null,
            ExcelLibExportPropertyParameter<T, XLAlignmentHorizontalValues?> horizontalAlignment = null,
            ExcelLibExportPropertyParameter<T, XLAlignmentVerticalValues?> verticalAlignment = null,
            ExcelLibExportPropertyParameter<T, XLCellValues?> dataType = null)
        {
            if (relativeColumnPlace < 1) throw new ArgumentOutOfRangeException("relativeColumnPlace", string.Format("Invalid column value: {0}", relativeColumnPlace));

            this._relativeColumnPlace = relativeColumnPlace;
            this._valueConverter = valueConverter;
            this._header = header;
            this._backgroundColor = backgroundColor;
            this._horizontalAlignment = horizontalAlignment;
            this._verticalAlignment = verticalAlignment;
            this._dataType = dataType;
        }

        /// <summary></summary>
        /// <param name="value"></param>
        /// <returns></returns>
        internal Color? GetBackgroundColor(T value)
        {
            return this._backgroundColor == null ? (Color?)null : this._backgroundColor.GetParameter(value);
        }
        /// <summary></summary>
        /// <param name="value"></param>
        /// <returns></returns>
        internal XLAlignmentHorizontalValues? GetHorizontalAlignment(T value)
        {
            return this._horizontalAlignment == null ? (XLAlignmentHorizontalValues?)null : this._horizontalAlignment.GetParameter(value);
        }        /// <summary></summary>
        /// <param name="value"></param>
        /// <returns></returns>
        internal XLAlignmentVerticalValues? GetVerticalAlignment(T value)
        {
            return this._verticalAlignment == null ? (XLAlignmentVerticalValues?)null : this._verticalAlignment.GetParameter(value);
        }
        /// <summary></summary>
        /// <param name="value"></param>
        /// <returns></returns>
        internal XLCellValues? GetDataType(T value)
        {
            return this._dataType == null ? (XLCellValues?)null : this._dataType.GetParameter(value);
        }

        /// <summary></summary>
        internal int RelativeColumnPlace { get { return this._relativeColumnPlace; } }
        /// <summary></summary>
        internal Func<T, string> ValueConverter { get { return this._valueConverter; } }
        /// <summary></summary>
        internal ExcelLibExportPropertyHeader Header { get { return this._header; } }
    }
}