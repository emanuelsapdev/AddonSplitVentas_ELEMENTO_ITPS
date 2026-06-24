using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Addon_AutoDivSalesOrd.Services
{
    public class ConverterService
    {

        public static decimal GetDecimalFromStringSAP(string val)
        {
            if (string.IsNullOrEmpty(val)) return 0;
            decimal result = decimal.Parse(val.Trim(), CultureInfo.InvariantCulture);
            return decimal.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public static DateTime GetDateTimeFromStringSAP(string value)
        {
            return DateTime.ParseExact(value.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
        }

        public static decimal GetDecimalFromCurrencyStringSAP(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return 0m;

            string cleaned = value
                .Replace("ARS", "")
                .Replace("USD", "")
                .Replace("\"", "")
                .Replace(" ", "")
                .Replace(".", "")
                .Replace(",", ".");

            return decimal.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result)
                ? result
                : 0m;
        }
    }
}
