using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Excel.IO
{
    public static class ExtensionMethods
    {
        public static string ReplaceDecimalSeparator(this string text)
        {
            return text.Replace('.', Convert.ToChar(CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)).Replace(',', Convert.ToChar(CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator));
        }
    }
}
