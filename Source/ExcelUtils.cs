using System;

using OfficeOpenXml;

namespace ExcelGridDataProviderEPPlus
{
    /// <summary>
    /// Useful static Simio EPPlus utility methods for 
    /// </summary>
    public static class ExcelUtils
    {
        /// <summary>
        /// Get the cell value as a string
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public static string GetCellValue(ExcelRange cell)
        {

            switch (cell.Value)
            {
                case double dd:
                    {
                        DateTime dt = DateTime.MinValue;
                        if (DateTime.TryParse(cell.Text, out dt))
                            return GetDateTimeOrNumericValueAsString(cell);
                        else
                            return cell.Text;
                    }
                    break;

                case string ss:
                        return ss;

                case DateTime dt:
                        return GetDateTimeOrNumericValueAsString(cell);

                case Boolean bb:
                    return cell.Value.ToString();

                default:
                    break;
            }



            ////    case CellValueType.Text:
            ////        {

            ////            return cell.Value.ToString();
            ////        }
            ////    case CellValueType.Numeric:
            ////        return GetDateTimeOrNumericValueAsString(cell);
            ////    case CellValueType.Boolean:
            ////        return cell.GetValue<bool>().ToString(System.Globalization.CultureInfo.InvariantCulture);
            ////    case CellValueType.None:
            ////        return null;
            ////    case CellValueType.Error:
            ////        return cell.Value.ToString();
            ////    case CellValueType.DateTime:
            ////        return GetDateTimeOrNumericValueAsString(cell);
            ////    case CellValueType.Unknown:
            ////        return "Unknown";
            ////}

            return null;
        }

        /// <summary>
        /// Returns the value of a single cell.
        /// If the type of the Cell's Value is DateTime, then return
        /// the value as DateTime, with adjustments to round it to seconds
        /// when the milliseconds is 995 or greater.
        /// If the type is Numeric, the return the Cell's value.
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static string GetDateTimeOrNumericValueAsString(ExcelRange cell)
        {

            if (cell.Value is DateTime)
            {
                DateTime dt = (DateTime) cell.Value;
                if (dt.Millisecond >= 995)
                {
                    // Excel stores things as Days from Jan 1, 1900 (or Jan 1, 1904 for the Mac)
                    // This can (apparently) result in some values like
                    // 1/7/2016 4:29:59.999 when what was in excel was shown as 1/7/2016 4:30:00, so....
                    // If we are very, very close to the next second, so we'll go to the next second, since 
                    //  the ToString() will simply strip off any sub-second values.
                    dt = dt.AddSeconds(1.0);
                }
                return dt.ToString(); // Simio will first try to parse dates in the current culture
            }
            else if (cell.Value is Decimal dd)
            {
                return dd.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
            else
            {
                return cell.Value.ToString();
            }
        }

        /// <summary>
        /// Get a cell as text and try and parse it as a decimal
        /// Return false if the value is null or isn't a decimal, in which case the dd argument is untouched.
        /// Return true if a legitimate double (dd) is found and set.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="dd"></param>
        /// <returns></returns>
        public static Decimal? GetCellAsDecimal(ExcelRange cell)
        {
            if (cell?.Value == null)
                return null;

            if (Decimal.TryParse(cell.Text, out decimal newValue))
                return newValue;
            else
                return null;

        }

    } // class
}


