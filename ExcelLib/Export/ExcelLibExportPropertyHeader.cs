using System;

using ClosedXML.Excel;

namespace ExcelLib.Export
{
    /// <summary></summary>
    /// <typeparam name="T"></typeparam>
    public sealed class ExcelLibExportPropertyHeader
    {
        /// <summary></summary>
        private readonly string _header;
        /// <summary></summary>
        private readonly XLAlignmentHorizontalValues? _horizontalAlignment;
        /// <summary></summary>
        private readonly XLAlignmentVerticalValues? _verticalAlignment;

        /// <summary></summary>
        /// <param name="header"></param>
        /// <param name="horizontalAlignment"></param>
        /// <param name="verticalAlignment"></param>
        public ExcelLibExportPropertyHeader(string header, XLAlignmentHorizontalValues? horizontalAlignment = null, XLAlignmentVerticalValues? verticalAlignment = null)
        {
            if (header == null) throw new ArgumentNullException("header");

            this._header = header;
            this._horizontalAlignment = horizontalAlignment;
            this._verticalAlignment = verticalAlignment;
        }

        /// <summary></summary>
        internal string Header { get { return this._header; } }
        /// <summary></summary>
        internal XLAlignmentHorizontalValues? HorizontalAlignment { get { return this._horizontalAlignment; } }
        /// <summary></summary>
        internal XLAlignmentVerticalValues? VerticalAlignment { get { return this._verticalAlignment; } }
    }
}