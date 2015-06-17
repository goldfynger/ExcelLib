using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

using ClosedXML.Excel;

using ExcelLib.Export;

namespace _Run_ExcelLib
{
    /// <summary></summary>
    class Program
    {
        /// <summary></summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            var list = new List<ExcelTuple>
            {
                new ExcelTuple { BooleanField = false, IntegerField = 1, StringField = "One" },
                new ExcelTuple { BooleanField = true, IntegerField = 2, StringField = "Two" },
                new ExcelTuple { BooleanField = false, IntegerField = 3, StringField = "Three" },
                new ExcelTuple { BooleanField = true, IntegerField = 4, StringField = "Four" },
            };

            var booleanProperty = new ExcelLibExportProperty<bool>(2,
                header: new ExcelLibExportPropertyHeader("Флаг"),
                dataType: new ExcelLibExportPropertyParameter<bool, XLCellValues?>(XLCellValues.Boolean),
                verticalAlignment: new ExcelLibExportPropertyParameter<bool, XLAlignmentVerticalValues?>(XLAlignmentVerticalValues.Center),
                horizontalAlignment: new ExcelLibExportPropertyParameter<bool, XLAlignmentHorizontalValues?>(XLAlignmentHorizontalValues.Right));
            var integerProperty = new ExcelLibExportProperty<int>(4,
                header: new ExcelLibExportPropertyHeader("Значение"),
                verticalAlignment: new ExcelLibExportPropertyParameter<int, XLAlignmentVerticalValues?>(XLAlignmentVerticalValues.Top),
                dataType: new ExcelLibExportPropertyParameter<int, XLCellValues?>(XLCellValues.Number));
            var stringProperty = new ExcelLibExportProperty<string>(6,
                header: new ExcelLibExportPropertyHeader("Описание"),
                verticalAlignment: new ExcelLibExportPropertyParameter<string, XLAlignmentVerticalValues?>(XLAlignmentVerticalValues.Bottom),
                backgroundColor: new ExcelLibExportPropertyParameter<string, Color?>(Color.Green));

            Func<ExcelTuple, List<ExcelLibExportValue>> func = t =>
                new List<ExcelLibExportValue>
                {
                    new ExcelLibExportValue<bool>(t.BooleanField, booleanProperty),
                    new ExcelLibExportValue<int>(t.IntegerField, integerProperty),
                    new ExcelLibExportValue<string>(t.StringField, stringProperty)
                };

            var excelLibExportRange = new ExcelLibExportRange<ExcelTuple>(list, func);

            var excelLibExportString = new ExcelLibExportString("Строка", 6,
                horizontalAlignment: XLAlignmentHorizontalValues.Center,
                verticalAlignment: XLAlignmentVerticalValues.Bottom,
                backgroundColor: Color.Red);

            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Page one");

            var resultRange = excelLibExportRange.FillValues(ws, 6, 1);

            resultRange.SetAutoFilter();

            excelLibExportString.FillString(ws, 3, 1);

            ws.Columns().AdjustToContents();

            using (var stream = File.Open(@"..\..\..\TestFiles\Result.xlsx", FileMode.Create)) wb.SaveAs(stream);
        }

        /// <summary></summary>
        public sealed class ExcelTuple
        {
            /// <summary></summary>
            public string StringField { get; set; }
            /// <summary></summary>
            public int IntegerField { get; set; }
            /// <summary></summary>
            public bool BooleanField { get; set; }
        }
    }
}