using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDnaUdf
{
    internal class DateTimeLunar
    {

        #region 农历常量

        /// <summary>
        /// 公历与农历互转，公历最小时间
        /// </summary>
        private static DateTime MinDay = new DateTime(1900, 1, 31);

        /// <summary>
        /// 公历与农历互转，公历最大时间
        /// </summary>
        private static DateTime MaxDay = new DateTime(2100, 12, 31);

        /// <summary>
        /// 农历月份字符：正、二、三……十、冬、腊
        /// </summary>
        private static string MonthStr = "正贰叁肆伍陆柒捌九拾冬腊";

        /// <summary>
        /// 农历日子字符，初、十、廿、卅
        /// </summary>
        private static string NumeralsChinese2 = "初十廿卅";
        private static string NumeralsChinese = "〇一二三四五六七八九";

        private static string zhiStr = "子丑寅卯辰巳午未申酉戌亥";
        private static string animalStr = "鼠牛虎兔龙蛇马羊猴鸡狗猪";

        //private static string nStr2 = "初十廿卅";

        /// <summary>
        /// 农历数据
        /// </summary>
        private static int[] LunarDateArray = new int[] {
            0x84B6BF,
     0x04AE53,0x0A5748,0x5526BD,0x0D2650,0x0D9544,
     0x46AAB9,0x056A4D,0x09AD42,0x24AEB6,0x04AE4A, //1901-1910

     0x6A4DBE,0x0A4D52,0x0D2546,0x5D52BA,0x0B544E,
     0x0D6A43,0x296D37,0x095B4B,0x749BC1,0x049754, //1911-1920

     0x0A4B48,0x5B25BC,0x06A550,0x06D445,0x4ADAB8,
     0x02B64D,0x095742,0x2497B7,0x04974A,0x664B3E, //1921-1930

     0x0D4A51,0x0EA546,0x56D4BA,0x05AD4E,0x02B644,
     0x393738,0x092E4B,0x7C96BF,0x0C9553,0x0D4A48, //1931-1940

     0x6DA53B,0x0B554F,0x056A45,0x4AADB9,0x025D4D,
     0x092D42,0x2C95B6,0x0A954A,0x7B4ABD,0x06CA51, //1941-1950

     0x0B5546,0x555ABB,0x04DA4E,0x0A5B43,0x352BB8,
     0x052B4C,0x8A953F,0x0E9552,0x06AA48,0x7AD53C, //1951-1960

     0x0AB54F,0x04B645,0x4A5739,0x0A574D,0x052642,
     0x3E9335,0x0D9549,0x75AABE,0x056A51,0x096D46, //1961-1970

     0x54AEBB,0x04AD4F,0x0A4D43,0x4D26B7,0x0D254B,
     0x8D52BF,0x0B5452,0x0B6A47,0x696D3C,0x095B50, //1971-1980

     0x049B45,0x4A4BB9,0x0A4B4D,0xAB25C2,0x06A554,
     0x06D449,0x6ADA3D,0x0AB651,0x093746,0x5497BB, //1981-1990

     0x04974F,0x064B44,0x36A537,0x0EA54A,0x86B2BF,
     0x05AC53,0x0AB647,0x5936BC,0x092E50,0x0C9645, //1991-2000

     0x4D4AB8,0x0D4A4C,0x0DA541,0x25AAB6,0x056A49,
     0x7AADBD,0x025D52,0x092D47,0x5C95BA,0x0A954E, //2001-2010

     0x0B4A43,0x4B5537,0x0AD54A,0x955ABF,0x04BA53,
     0x0A5B48,0x652BBC,0x052B50,0x0A9345,0x474AB9, //2011-2020

     0x06AA4C,0x0AD541,0x24DAB6,0x04B64A,0x69573D,
     0x0A4E51,0x0D2646,0x5E933A,0x0D534D,0x05AA43, //2021-2030

     0x36B537,0x096D4B,0xB4AEBF,0x04AD53,0x0A4D48,
     0x6D25BC,0x0D254F,0x0D5244,0x5DAA38,0x0B5A4C, //2031-2040

     0x056D41,0x24ADB6,0x049B4A,0x7A4BBE,0x0A4B51,
     0x0AA546,0x5B52BA,0x06D24E,0x0ADA42,0x355B37, //2041-2050

     0x09374B,0x8497C1,0x049753,0x064B48,0x66A53C,
     0x0EA54F,0x06B244,0x4AB638,0x0AAE4C,0x092E42, //2051-2060

     0x3C9735,0x0C9649,0x7D4ABD,0x0D4A51,0x0DA545,
     0x55AABA,0x056A4E,0x0A6D43,0x452EB7,0x052D4B, //2061-2070

     0x8A95BF,0x0A9553,0x0B4A47,0x6B553B,0x0AD54F,
     0x055A45,0x4A5D38,0x0A5B4C,0x052B42,0x3A93B6, //2071-2080

     0x069349,0x7729BD,0x06AA51,0x0AD546,0x54DABA,
     0x04B64E,0x0A5743,0x452738,0x0D264A,0x8E933E, //2081-2090

     0x0D5252,0x0DAA47,0x66B53B,0x056D4F,0x04AE45,
     0x4A4EB9,0x0A4D4C,0x0D1541,0x2D92B5 ,0x0D5249,//2091-2100
        };

        #endregion 农历常量

        /// <summary>
        /// 公历每月第一天是公历年中的第几天（非润年）
        /// </summary>
        private static int[] SolarNormalYday = { 1, 32, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335 };

        /// <summary>
        /// 公历每月第一天是公历年中的第几天（润年）
        /// </summary>
        private static int[] SolarLeapYday = { 1, 32, 61, 92, 122, 153, 183, 214, 245, 275, 306, 336 };

        #region 私有变量

        /// <summary>
        /// 私有变量，能否公历转农历
        /// </summary>
        private bool _Solar2Lunar = false;

        /// <summary>
        /// 私有变量，该年的农历数据
        /// </summary>
        private int _LunarData;

        /// <summary>
        /// 私有变量，公历日期
        /// </summary>
        private DateTime _SolarDate;

        /// <summary>
        /// 私有变量，农历年份
        /// </summary>
        private int _LunarYear;
        /// <summary>
        /// 私有变量，农历年份中文
        /// </summary>
        private string _LunarYearChinese;

        /// <summary>
        /// 私有变量，农历月份1~13
        /// </summary>
        private int _LunarMonth;

        /// <summary>
        /// 私有变量，农历实际月份1~12
        /// </summary>
        private int _LunarMonth2;
        /// <summary>
        /// 私有变量，农历日中文
        /// </summary>
        private string _LunarMonth2Chinese;
        /// <summary>
        /// 私有变量，农历日
        /// </summary>
        private int _LunarDay;
        /// <summary>
        /// 私有变量，农历日中文
        /// </summary>
        private string _LunarDayChinese;
        /// <summary>
        /// 私有变量，公历年份
        /// </summary>
        private int _SolarYear;

        /// <summary>
        /// 私有变量，公历月份
        /// </summary>
        private int _SolarMonth;

        /// <summary>
        /// 私有变量，公历日
        /// </summary>
        private int _SolarDay;

        /// <summary>
        /// 私有变量，农历年是否闰年
        /// </summary>
        private bool _IsLeapLunarYear;
        /// <summary>
        /// 私有变量，该月是否农历润月
        /// </summary>
        private bool _IsLeapLunarMonth;
        /// <summary>
        /// 闰月字符串
        /// </summary>
        private string _LeapLunarMonthChinese = "润";

        /// <summary>
        /// 私有变量，农历年哪个月是润月，6代表润六月，0表示该年无润月
        /// </summary>
        private int _WhereLeapLunarMonth;

        /// <summary>
        /// 私有变量，农历年1~13个月每个月的天数，闰月之后的月份数+1
        /// </summary>
        private int[] _DaysOfLunarMonth = { 29, 29, 29, 29, 29, 29, 29, 29, 29, 29, 29, 29, 0 };

        /// <summary>
        /// 私有变量，公历日期是农历年的第几天
        /// </summary>
        private int _DayOfLunarYear;

        #endregion 私有变量

        #region 公共变量

        /// <summary>
        /// 公历能否转农历
        /// </summary>
        public bool Solar2Lunar
        {
            get
            {
                return _Solar2Lunar;
            }
        }

        /// <summary>
        /// 公历日期
        /// </summary>
        public DateTime SolarDate
        {
            set
            {
                _SolarDate = value;
                GetLunar();
            }
            get
            {
                return _SolarDate;
            }
        }

        /// <summary>
        /// 农历年份
        /// </summary>
        public int LunarYear
        {
            get
            {
                return _LunarYear;
            }
        }

        /// <summary>
        /// 农历月份
        /// </summary>
        public int LunarMonth
        {
            get
            {
                return _LunarMonth;
            }
        }

        /// <summary>
        /// 农历日
        /// </summary>
        public int LunarDay
        {
            get
            {
                return _LunarDay;
            }
        }

        /// <summary>
        /// 农历年是否闰年
        /// </summary>
        public bool IsLeapLunarYear
        {
            get
            {
                return _IsLeapLunarYear;
            }
        }

        /// <summary>
        /// 农历年润哪个月，6代表润六月
        /// </summary>
        public int WhereLeapLunarMonth
        {
            get
            {
                if (_WhereLeapLunarMonth > 0)
                {
                    return _WhereLeapLunarMonth - 1;
                }
                else
                {
                    return 0;
                }
            }
        }

        #endregion 公共变量

        #region 构造方法

        public DateTimeLunar()
        {
        }
        /// <summary>
        /// 使用日期来构造
        /// </summary>
        /// <param name="datetime"></param>
        public DateTimeLunar(DateTime datetime)
        {
            this.SolarDate = datetime;
            GetLunar();
        }
        /// <summary>
        /// 使用8位数日期字符串来构造
        /// </summary>
        /// <param name="datetimeStr"></param>
        public DateTimeLunar(string datetimeStr)
        {
            DateTime dt = DateTime.ParseExact(datetimeStr, "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None);
            this.SolarDate = dt;
            GetLunar();
        }
        #endregion 构造方法

        #region 私有方法

        /// <summary>
        /// 获取指定年份对应的农历数据
        /// </summary>
        /// <param name="year"></param>
        /// <returns></returns>
        private int GetLunarData(int year)
        {
            return LunarDateArray[year - MinDay.Year];
        }

        /// <summary>
        /// 获取农历相关信息
        /// </summary>
        private void GetLunar()
        {
            if (_SolarDate >= MinDay && _SolarDate <= MaxDay)
            {
                _Solar2Lunar = true;
            }
            else
            {
                _Solar2Lunar = false;
            }
            if (!_Solar2Lunar)
            {
                return;
            }
            _SolarYear = _SolarDate.Year;
            _LunarYear = SolarDateToLunarYear(_SolarDate);
            _LunarData = LunarDateArray[_LunarYear - MinDay.Year];


            DateTime dt2 = GetLunarFirstDayToSolarDate(_SolarDate.Year);//该年春节日期

            if (_SolarDate < dt2)//农历年份比公历少1
            {
                DateTime dt1 = GetLunarFirstDayToSolarDate(_SolarDate.Year - 1);//前一年春节日期
                _DayOfLunarYear = DaysBetweenTwoSolarDates(dt1, _SolarDate);
            }
            else
            {
                _DayOfLunarYear = DaysBetweenTwoSolarDates(dt2, _SolarDate);
            }

            _LunarMonth = 1;

            for (; _LunarMonth <= 13; _LunarMonth++)
            {
                int daysOfLunarMonth = GetDaysInLunarMonth(_LunarYear, _LunarMonth);
                if (_DayOfLunarYear <= daysOfLunarMonth)
                {
                    break;
                }
                else
                {
                    _DayOfLunarYear -= daysOfLunarMonth;
                }
            }
            _LunarDay = _DayOfLunarYear + 1;

            //处理润月
            _WhereLeapLunarMonth = (_LunarData >> 20) & 0xf;
            //否润年
            if (_WhereLeapLunarMonth > 0)
            {
                _IsLeapLunarYear = true;
                //是否润月
                if ((_WhereLeapLunarMonth + 1) == _LunarMonth)
                {
                    _IsLeapLunarMonth = true;
                }
                else
                {
                    _IsLeapLunarMonth = false;
                }
                //润年时，实际月份
                if (_LunarMonth > _WhereLeapLunarMonth)
                {
                    _LunarMonth2 = _LunarMonth - 1;
                }
                else
                {
                    _LunarMonth2 = _LunarMonth;
                }
            }
            else
            {
                _IsLeapLunarYear = false;
                _LunarMonth2 = _LunarMonth;
            }
        }
        /// <summary>
        /// 获取农历月、日中文表示文字
        /// </summary>
        private void GetLunarMonthAndDayChinese()
        {
            string daystr = null;
            if (_LunarDay == 10 | _LunarDay == 20 | _LunarDay == 30)
            {
                daystr = "十";
            }
            else
            {
                daystr = NumeralsChinese[_LunarDay % 10].ToString();
            }
            _LunarDayChinese = NumeralsChinese2[(int)_LunarDay / 10].ToString() + daystr;

            if (_IsLeapLunarMonth)
            {
                _LeapLunarMonthChinese = "润";
            }
            else
            {
                _LeapLunarMonthChinese = null;
            }

            _LunarMonth2Chinese = MonthStr[_LunarMonth2 - 1].ToString();


        }

        #endregion 私有方法

        #region 公共方法
        /// <summary>
        /// 获取公历对应的农历
        /// </summary>
        /// <param name="LunarString"></param>
        /// <returns></returns>
        public string GetLunarDate(LunarString LunarString)
        {
            string s = null;
            switch (LunarString)
            {
                case LunarString.农历润六月初八:
                    GetLunarMonthAndDayChinese();
                    s = string.Format("农历 {0}月{1}", _LeapLunarMonthChinese + _LunarMonth2Chinese, _LunarDayChinese);
                    break;
                case LunarString.农历润六月:
                    GetLunarMonthAndDayChinese();
                    if (_LunarDay == 1)
                    {
                        s = string.Format("农历 {0}月", _LeapLunarMonthChinese + _LunarMonth2Chinese);
                    }
                    else
                    {
                        s = string.Format("农历 {0}月{1}", _LeapLunarMonthChinese + _LunarMonth2Chinese, _LunarDayChinese);
                    }
                    break;
                case LunarString.农历二零一七年润六月初八:
                    for (int i = 0; i < _LunarYear.ToString().Length; i++)
                    {
                        _LunarYearChinese = _LunarYearChinese + NumeralsChinese.Substring(int.Parse(_LunarYear.ToString().Substring(i, 1)), 1);
                    }

                    GetLunarMonthAndDayChinese();
                    s = string.Format("农历 {0}年{1}月{2}", _LunarYearChinese, _LeapLunarMonthChinese + _LunarMonth2Chinese, _LunarDayChinese);
                    break;
                case LunarString.农历润6月8日:
                    s = string.Format("农历 {0}月{1}日", _LunarYearChinese + _LunarMonth2, _LunarDay);
                    break;
                default:
                    break;
            }
            return s;
        }
        /// <summary>
        /// 获取公历日期
        /// </summary>
        /// <param name="lunarYear"></param>
        /// <param name="lunarMonth"></param>
        /// <param name="lunarDay"></param>
        /// <param name="theMonthIsLeap"></param>
        /// <returns></returns>
        public DateTime GetSolarDate(int lunarYear, int lunarMonth, int lunarDay, bool theMonthIsLeap)
        {
            DateTime dt = GetLunarFirstDayToSolarDate(lunarYear);//该年春节对应的公历

            int d = DaysBetweenTwoSolarDates(new DateTime(lunarYear, 1, 1), dt);//该年元旦距离春节的天数
            //农历日期距离该年春节的天数
            int days = DaysOfLunarFirstDayToThisLunarDate(lunarYear, lunarMonth, lunarDay, theMonthIsLeap);
            // int days = DaysOfSolarFirstDayToThisLunarDate(lunarYear, lunarMonth, lunarDay, theMonthIsLeap);
            DateTime dt1 = new DateTime(lunarYear, 1, 1);
            days = days + d;
            DateTime dt2 = dt1.AddDays(days);

            return dt2;
        }
        public int GetDaysInLunarMonth(int lunarYear, int lunarMonth)
        {
            int lunarData = GetLunarData(lunarYear);
            //获得该年润月份
            int leapMonth = (lunarData >> 20) & 0xf;
            if (leapMonth == 0 && lunarMonth == 13)
            {
                return 0;
            }
            int daysOfLunarMonth = 29;
            if (((lunarData >> (20 - lunarMonth)) & 0x1) == 1)//大月就30天
            {
                daysOfLunarMonth = 30;
            }
            return daysOfLunarMonth;
        }
        /// <summary>
        /// 返回指定年份春节所对应的公历日期
        /// </summary>
        /// <param name="year">年份</param>
        /// <returns></returns>
        public DateTime GetLunarFirstDayToSolarDate(int year)
        {
            int lunarData = GetLunarData(year);
            int month;
            if (((lunarData & 0x0060) >> 5) == 1)//0x0060 转二进制为：1100000
            {
                month = 1;
            }
            else
            {
                month = 2;
            }
            int day = lunarData & 0x1F;

            DateTime tDateTime = new DateTime(year, month, day);
            return tDateTime;
        }

        /// <summary>
        /// 返回公历日期对应的农历年份
        /// </summary>
        /// <param name="dt">公历日期</param>
        /// <returns></returns>
        public int SolarDateToLunarYear(DateTime dt)
        {
            DateTime dtSpringFestival = GetLunarFirstDayToSolarDate(dt.Year);
            int year = dt.Year;
            if (dt < dtSpringFestival)
            {
                return year - 1;
            }
            else
            {
                return year;
            }
        }
        /// <summary>
        /// 返回指定公历年份是否为公历闰年
        /// </summary>
        /// <param name="year">公历年份</param>
        /// <returns></returns>
        public bool LeapSolarYear(int year)
        {
            if (year % 400 == 0)
            {
                return true;
            }
            else if (year % 4 == 0 && year % 100 != 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// 计算两个公历日期之间的天数，同一天为0
        /// </summary>
        /// <param name="minDt"></param>
        /// <param name="maxDt"></param>
        /// <returns></returns>
        public int DaysBetweenTwoSolarDates(DateTime minDt, DateTime maxDt)
        {
            TimeSpan ts = maxDt.Subtract(minDt);
            int days = ts.Days;
            return days;
        }
        /// <summary>
        /// 指定农历日期距离当年春节的天数，春节这一天为0
        /// </summary>
        /// <param name="lunarYear">农历年</param>
        /// <param name="lunarMonth">农历月</param>
        /// <param name="lunarDay">农历日</param>
        /// <param name="theMonthIsLeap">该月（lunarMonth）是否闰月</param>
        /// <returns></returns>
        public int DaysOfLunarFirstDayToThisLunarDate(int lunarYear, int lunarMonth, int lunarDay, bool theMonthIsLeap)
        {
            int lunarData = LunarDateArray[lunarYear - MinDay.Year];
            //获得该年润月份
            int leapMonth = (lunarData >> 20) & 0xf;
            //月份转换为13月制
            if (leapMonth > 0 && lunarMonth > leapMonth)
            {
                lunarMonth++;
            }
            else if (lunarMonth == leapMonth && theMonthIsLeap == true)
            {
                lunarMonth++;
            }

            int days = 0;
            for (int m = 1; m < lunarMonth; m++)
            {
                int daysOfLunarMonth = GetDaysInLunarMonth(lunarYear, m);
                days = days + daysOfLunarMonth;
            }
            days = days + lunarDay - 1;

            return days;
        }
        #endregion 公共方法

    }

    /// <summary>
    /// 农历日期格式化显示
    /// </summary>
    public enum LunarString
    {
        /// <summary>
        /// 农历月、日
        /// </summary>
        农历润六月初八,
        /// <summary>
        /// 农历月、日。若逢初一则省略日
        /// </summary>
        农历润六月,
        /// <summary>
        /// 农历年、月、日
        /// </summary>
        农历二零一七年润六月初八,
        /// <summary>
        /// 农历月、日（阿拉伯数字）
        /// </summary>
        农历润6月8日,
        /// <summary>
        /// 农历年、月、日（阿拉伯数字）
        /// </summary>
        农历2017年润6月8日,

    }

    struct LunarTime
    {
        int lunarYear;
        int lunarMonth;
        int lunarDay;


    }
}
