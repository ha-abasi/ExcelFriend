using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFriend
{
    public class Utility
    {
        public static string getShamsiToMiladiString(string str)
        {
            var dt = getShamsiToMiladi(str);
            if (dt == null)
            {
                return "";
            }
            else
            {
                return dt?.ToString("yyyy/MM/dd");
            }
            
        }
        public static DateTime? getShamsiToMiladi(string str)
        {
            CultureInfo persianCulture = new CultureInfo("fa-IR");
            try
            {
                DateTime persianDateTime = DateTime.ParseExact(str, "yyyy/MM/dd", persianCulture);
                return persianDateTime;
                //sheet.Cell(index, destColumn).Value = persianDateTime.ToString("yyyy/MM/dd");
            }
            catch (Exception ex)
            {
                //ignore
            }


            return null;
        }
    }
}
