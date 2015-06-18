using System;

namespace ExcelLib
{
    /// <summary></summary>
    /// <typeparam name="T"></typeparam>
    public sealed class ExcelLibResultContainer<T>
    {
        /// <summary></summary>
        private readonly bool _isValid;
        /// <summary></summary>
        private readonly T _value;

        /// <summary></summary>
        /// <param name="isValid"></param>
        /// <param name="value"></param>
        private ExcelLibResultContainer(bool isValid, T value)
        {
            this._isValid = isValid;
            this._value = value;
        }

        /// <summary></summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static ExcelLibResultContainer<T> CreateValidValue(T value)
        {
            return new ExcelLibResultContainer<T>(true, value);
        }
        /// <summary></summary>
        /// <returns></returns>
        public static ExcelLibResultContainer<T> CreateInvalidValue()
        {
            return new ExcelLibResultContainer<T>(false, default(T));
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