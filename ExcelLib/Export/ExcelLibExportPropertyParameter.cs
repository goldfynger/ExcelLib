using System;

namespace ExcelLib.Export
{
    /// <summary></summary>
    /// <typeparam name="T"></typeparam>
    /// <typeparam name="TParam"></typeparam>
    public sealed class ExcelLibExportPropertyParameter<T, TParam>
    {
        /// <summary></summary>
        private bool _isInitialized = false;

        /// <summary></summary>
        private readonly TParam _parameter;

        /// <summary></summary>
        private readonly Func<T, TParam> _func;

        /// <summary></summary>
        public ExcelLibExportPropertyParameter(TParam parameter)
        {
            this._parameter = parameter;

            this._isInitialized = true;
        }

        /// <summary></summary>
        /// <param name="func"></param>
        public ExcelLibExportPropertyParameter(Func<T, TParam> func)
        {
            if (func == null) throw new ArgumentNullException("func");
            this._func = func;

            this._isInitialized = false;
        }

        /// <summary></summary>
        /// <param name="value"></param>
        /// <returns></returns>
        internal TParam GetParameter(T value)
        {
            return this._isInitialized ? this._parameter : this._func(value);
        }
    }
}