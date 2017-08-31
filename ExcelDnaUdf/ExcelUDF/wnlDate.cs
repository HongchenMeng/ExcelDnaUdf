using System;
using System.IO;
using System.Text;
using System.Linq;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace ExcelDnaUdf
{
    public partial class ExcelUDF
    {
        /// <summary>
        /// 获取农历
        /// </summary>
        /// <param name="dtStr">时间字符串</param>
        /// <returns></returns>
        [ExcelFunction(Category = "农历日期", IsMacroType = true, Description = "获取农历日期")]
        public static string GetLunarDate([ExcelArgument(Description = "时间")] string dtStr)
        {
            string[] dd = dtStr.Split('/');
            int y = int.Parse(dd[0]);
            int m = int.Parse(dd[1]);
            int d = int.Parse(dd[2]);
            DateTime dt = new DateTime(y, m, d);
            ChinaDate cd = new ChinaDate(dt);

            return cd.nlDate;
        }

        /// <summary>
        /// 获取农历干支
        /// </summary>
        /// <param name="dtStr">时间字符串</param>
        /// <returns></returns>
        [ExcelFunction(Category = "农历日期", IsMacroType = true, Description = "获取农历日期")]
        public static string GetGZ([ExcelArgument(Description = "时间")] string dtStr)
        {
            string[] dd = dtStr.Split('/');
            int y = int.Parse(dd[0]);
            int m = int.Parse(dd[1]);
            int d = int.Parse(dd[2]);
            DateTime dt = new DateTime(y, m, d);
            ChinaDate cd = new ChinaDate(dt);

            return cd.gzDate;
        }
        /// <summary>
        /// 获取公历日期
        /// </summary>
        /// <param name="lunarYear">农历年份</param>
        /// <param name="lunarMonth">农历月份</param>
        /// <param name="lunarDay">农历日</param>
        /// <param name="theMonthIsLeap">该月是否闰月</param>
        /// <returns></returns>
        [ExcelFunction(Category = "农历日期", IsMacroType = true, Description = "获取公历日期")]
        public static string GetSolarDate([ExcelArgument(Description = "农历年份")] int lunarYear, [ExcelArgument(Description = "农历月份")] int lunarMonth, [ExcelArgument(Description = "农历日")] int lunarDay, [ExcelArgument(Description = "该月是否农历润月")] bool theMonthIsLeap)
        {
            DateTimeLunar dl = new DateTimeLunar();
            DateTime dt = dl.GetSolarDate(lunarYear, lunarMonth, lunarDay, theMonthIsLeap);

            return dt.ToShortDateString();
        }
        /// <summary>
        /// 获取农历每月的天数
        /// </summary>
        /// <param name="lunarYear">农历年</param>
        /// <param name="lunarMonth">农历月</param>
        /// <returns></returns>
        [ExcelFunction(Category = "农历日期", IsMacroType = true, Description = "获取农历每月天数")]
        public static int GetDaysInLunarMonth([ExcelArgument(Description = "农历年份")] int lunarYear, [ExcelArgument(Description = "农历月份")] int lunarMonth)
        {
            DateTimeLunar dl = new DateTimeLunar();
            int days = dl.GetDaysInLunarMonth(lunarYear, lunarMonth);

            return days;
        }
    }
}
