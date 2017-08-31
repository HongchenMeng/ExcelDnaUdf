using System;
using SharpSxwnl;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDnaUdf
{
    class ChinaDate
    {
        public ChinaDate(DateTime dt)
        {
            getlun(dt);
        }

        private OB ob;
        /// <summary>
        /// 干支纪年 以立春为起始 丁酉年 丁未月 戊午日
        /// </summary>
        public string gzDate
        {
            get
            {
                return ob.Lyear2 + "年 " + ob.Lmonth2 + "月 " + ob.Lday2 + "日";
            }
        }
        /// <summary>
        /// 干支纪时
        /// </summary>
        public string gzTime
        {
            get
            {
                return ob.Ltime2.ToString();
            }
        }
        /// <summary>
        /// 农历月、日
        /// </summary>
        public string nlDate
        {
            get
            {
                return ob.Lleap + ob.Lmc + "月" + (ob.Ldn > 29 ? "大 " : "小 ") + ob.Ldc + "日";
            }
        }
        private void getlun(DateTime dt)
        {
            Lunar lun = new Lunar();
            double curJD, curTZ;
            sun_moon smc = new sun_moon();

            //DateTime nowDT = DateTime.Now;
            curTZ = -8;//TimeZone.CurrentTimeZone.GetUtcOffset(nowDT).Negate().TotalHours;     // 中国: 东 8 区
            curJD = LunarHelper.NowUTCmsSince19700101(dt) / 86400000d - 10957.5 - curTZ / 24d; //J2000起算的儒略日数(当前本地时间)
            JD.setFromJD(curJD + LunarHelper.J2000);
            string Cal_y = JD.Y.ToString();
            string Cal_m = JD.M.ToString();

            curJD = LunarHelper.int2(curJD + 0.5);

            // double By = LunarHelper.year2Ayear<string>(this.Cal_y.Text);
            //// C#: 注: 使用上句也可以, 如果在调用泛型方法时, 不指定类型, C# 编译器将自动推断其类型
            double By = LunarHelper.year2Ayear(Cal_y);    // 自动推断类型为: string
            double Bm = int.Parse(Cal_m);
            lun.yueLiHTML((int)By, (int)Bm, curJD, dt.Day);//html月历生成,结果返回在lun中,curJD为当前日期(用于设置今日标识)
                                                           //显示n指定的日期信息
            ob = lun.lun[dt.Day - 1];
        }
        #region 计算并返回日月升中降信息

        private string RTS1(double jd, double vJ, double vW, double tz)
        {
            SZJ.calcRTS(jd, 1, vJ, vW, tz); //升降计算,使用北时时间,tz=-8指东8区,jd+tz应在当地正午左右(误差数小时不要紧)
            string s;
            LunarInfoListT<double> ob = SZJ.rts[0];
            JD.setFromJD(jd + LunarHelper.J2000);
            s = "日出 " + ob.s + " 日落 " + ob.j + " 中天 " + ob.z + "\r\n"
                        + "月出 " + ob.Ms + " 月落 " + ob.Mj + " 月中 " + ob.Mz + "\r\n"
                        + "晨起天亮 " + ob.c + " 晚上天黑 " + ob.h + "\r\n"
                        + "日照时间 " + ob.sj + " 白天时间 " + ob.ch + "\r\n";
            return s;
        }

        #endregion 计算并返回日月升中降信息
    }
}
