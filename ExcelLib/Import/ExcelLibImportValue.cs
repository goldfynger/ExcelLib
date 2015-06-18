using System;

namespace ExcelLib.Import
{
    /// <summary></summary>
    public abstract class ExcelLibImportValue
    {
        /// <summary></summary>
        internal readonly Guid Guid;

        /// <summary></summary>
        /// <param name="guid"></param>
        protected ExcelLibImportValue(Guid guid)
        {
            this.Guid = guid;
        }
    }

    /// <summary></summary>
    /// <typeparam name="T"></typeparam>
    public sealed class ExcelLibImportValue<T> : ExcelLibImportValue
    {
        /// <summary></summary>
        private readonly bool _isValid;
        /// <summary></summary>
        private readonly T _value;

        /// <summary></summary>
        /// <param name="isValid"></param>
        /// <param name="value"></param>
        private ExcelLibImportValue(bool isValid, T value, Guid guid) : base(guid)
        {
            this._isValid = isValid;
            this._value = value;
        }

        /// <summary></summary>
        /// <param name="value"></param>
        /// <returns></returns>
        internal static ExcelLibImportValue<T> CreateValidValue(T value, Guid guid)
        {
            return new ExcelLibImportValue<T>(true, value, guid);
        }
        /// <summary></summary>
        /// <returns></returns>
        internal static ExcelLibImportValue<T> CreateInvalidValue(Guid guid)
        {
            return new ExcelLibImportValue<T>(false, default(T), guid);
        }

        /// <summary></summary>
        public bool IsValid { get { return this._isValid; } }

        /// <summary></summary>
        public T Value
        {
            get
            {
                if (!this._isValid) throw new InvalidOperationException("Value is invalid");

                return this._value;
            }
        }
    }
}