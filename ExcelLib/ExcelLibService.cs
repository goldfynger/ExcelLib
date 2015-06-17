using System;
using System.Linq;
using System.Text;

namespace ExcelLib
{
    /// <summary></summary>
    public static class ExcelLibService
    {
        /// <summary>Проверка на валидность всех символов строки для сериализации в XML, замена, если невалидные символы найдены.</summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string CheckAndChangeXmlString(string input)
        {
            if (!(input.Select(Convert.ToInt32).Any(i => !IsLegalXmlChar(i)))) return input;

            var buffer = new int[input.Length];

            for (var i = 0; i < input.Length; i++)
            {
                var integer = Convert.ToInt32(input[i]);
                buffer[i] = IsLegalXmlChar(integer) ? integer : 0x20;
            }

            var output = new StringBuilder();
            foreach (var t in buffer) output.Append((char)t);
            return output.ToString();
        }

        /// <summary>Возвращает false, если символ не валидный для сериализации в XML.</summary>
        /// <param name="character"></param>
        /// <returns></returns>
        private static bool IsLegalXmlChar(int character)
        {
            return
                (
                    character == 0x9 /* == '\t' == 9   */          ||
                    character == 0xA /* == '\n' == 10  */          ||
                    character == 0xD /* == '\r' == 13  */          ||
                    (character >= 0x20 && character <= 0xD7FF) ||
                    (character >= 0xE000 && character <= 0xFFFD) ||
                    (character >= 0x10000 && character <= 0x10FFFF)
                );
        }
    }
}