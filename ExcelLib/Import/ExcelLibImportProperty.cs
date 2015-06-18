using System;
using System.Diagnostics;

using ClosedXML.Excel;

namespace ExcelLib.Import
{
    /// <summary></summary>
    public abstract class ExcelLibImportProperty
    {
        /// <summary></summary>
        internal readonly Guid Guid = System.Guid.NewGuid();

        /// <summary></summary>
        /// <param name="worksheet"></param>
        /// <param name="row"></param>
        /// <param name="firstColumn"></param>
        /// <returns></returns>
        internal abstract ExcelLibImportValue ParseValue(IXLWorksheet worksheet, int row, int firstColumn);

        /// <summary></summary>
        internal abstract int RelativeColumnPlace { get; }
    }

    /// <summary></summary>
    /// <typeparam name="T"></typeparam>
    public sealed class ExcelLibImportProperty<T> : ExcelLibImportProperty
    {
        /// <summary></summary>
        private readonly int _relativeColumnPalce;
        /// <summary></summary>
        private readonly Func<object, ExcelLibResultContainer<T>> _convertValue;

        /// <summary></summary>
        /// <param name="relativeColumnPlace"></param>
        /// <param name="includeIfEmpty"></param>
        /// <param name="convertValue"></param>
        /// <param name="includeIfInvalid"></param>
        /// <param name="validateValue"></param>
        public ExcelLibImportProperty(int relativeColumnPlace, Func<object, ExcelLibResultContainer<T>> convertValue)
        {
            if (relativeColumnPlace < 1) throw new ArgumentOutOfRangeException("relativeColumnPlace", "Invaild column place");

            if (convertValue == null) throw new ArgumentNullException("convertValue");

            this._relativeColumnPalce = relativeColumnPlace;
            this._convertValue = convertValue;
        }

        /// <summary></summary>
        /// <param name="worksheet"></param>
        /// <param name="row"></param>
        /// <param name="firstColumn"></param>
        /// <returns></returns>
        internal override ExcelLibImportValue ParseValue(IXLWorksheet worksheet, int row, int firstColumn)
        {
            var cell = worksheet.Cell(row, firstColumn + this._relativeColumnPalce - 1);

            try
            {
                var result = this._convertValue(cell.Value);

                return result.IsValid ? ExcelLibImportValue<T>.CreateValidValue(result.Value, this.Guid) : ExcelLibImportValue<T>.CreateInvalidValue(this.Guid);
            }
            catch (Exception ex)
            {
                Debug.Fail(ex.Message);

                return ExcelLibImportValue<T>.CreateInvalidValue(this.Guid);
            }
        }

        /// <summary></summary>
        internal override int RelativeColumnPlace
        {
            get { return this._relativeColumnPalce; }
        }

        /// <summary></summary>
        /// <param name="container"></param>
        /// <returns></returns>
        public ExcelLibImportValue<T> GetValue(ExcelLibImportValuesContainer container)
        {
            return container.GetValue(this.Guid) as ExcelLibImportValue<T>;
        }
    }
}