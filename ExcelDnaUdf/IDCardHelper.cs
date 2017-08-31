using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace ExcelDnaUdf
{
    /// <summary>
    /// 身份证号码信息
    /// </summary>
    public class IDCardHelper
    {
        private string _id;
        private bool _checkBoll = true;
        private string _checkStr;//检查身份证号码是否有效
        private DateTime _birthday;
        private string _age;//年龄
        private string _sex;
        private string _area;//身份证前6位地址
        #region 公有属性
        /// <summary>
        /// 身份证号码
        /// </summary>
        public string ID
        {
            get
            {
                return _id;
            }
            set
            {
                _id = value;
                checkid();
            }
        }
        /// <summary>
        /// 身份证号码是否有效
        /// </summary>
        public bool checkBoll
        {
            get
            {
                return _checkBoll;
            }
        }
        /// <summary>
        /// 无效身份证号码的原因
        /// </summary>
        public string CheckStr
        {
            get
            {
                return _checkStr;
            }
        }
        /// <summary>
        /// 出生日期
        /// </summary>
        public DateTime Birthday
        {
            get
            {
                return _birthday;
            }
        }
        /// <summary>
        /// 身份证上的年龄，5岁以下有月，1岁以下有天
        /// </summary>
        public string Age
        {
            get
            {
                getAge();
                return _age;
            }
        }
        /// <summary>
        /// 性别
        /// </summary>
        public string Sex
        {
            get
            {
                getSex();
                return _sex;
            }
        }
        /// <summary>
        /// 身份证前6位地址
        /// </summary>
        public string Area
        {
            get
            {
                getArea();
                return _area;
            }
        }
        #endregion
        /// <summary>
        /// 构造函数
        /// </summary>
        public IDCardHelper()
        {

        }
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="ID">身份证号码</param>
        public IDCardHelper(string ID)
        {
            _id = ID;
            checkid();
        }
        #region 私有方法
        private void checkid()
        {
            if (!Regex.IsMatch(_id, @"^(^\d{15}$|^\d{18}$|^\d{17}(\d|X|x))$"))
            {
                _checkBoll = false;
                _checkStr = "身份证号码位应为15或18位";
            }
            else if (_id.Length == 15)//15位判断
            {
                string bd = "19" + _id.Substring(6, 6);
                if (!DateTime.TryParseExact(bd, "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None, out _birthday))
                {
                    _checkBoll = false;
                    _checkStr = "出生日期错误";
                }
            }
            else//18位判断
            {
                string bd = _id.Substring(6, 8);
                if (!DateTime.TryParseExact(bd, "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None, out _birthday))
                {
                    _checkBoll = false;
                    _checkStr = "出生日期错误";
                }
                else//检查识别码
                {
                    if (id18() != _id.Substring(17, 1).ToLower())
                    {
                        _checkBoll = false;
                        _checkStr = "最后一位应是：" + id18();
                    }
                }
            }
        }
        private string id18()
        {
            //1、将前面的身份证号码17位数分别乘以不同的系数。从第一位到第十七位的系数分别为：7－9－10－5－8－4－2－1－6－3－7－9－10－5－8－4－2。
            //2、将这17位数字和系数相乘的结果相加。
            //3、用加出来和除以11，看余数是多少？
            //4、余数只可能有0－1－2－3－4－5－6－7－8－9－10这11个数字。其分别对应的最后一位身份证的号码为1－0－X －9－8－7－6－5－4－3－2。(即馀数0对应1，馀数1对应0，馀数2对应X...)
            //5、通过上面得知如果余数是3，就会在身份证的第18位数字上出现的是9。如果对应的数字是2，身份证的最后一位号码就是罗马数字x。


            char[] id17 = _id.Remove(17).ToCharArray();//身份证前17位
            string[] id17Coefficient = ("7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2").Split(',');//第1~17位对应的系数
            string[] remainderCode = ("1,0,x,9,8,7,6,5,4,3,2").Split(',');//余数0~10对应的识别码

            int sum = 0;
            for (int i = 0; i < 17; i++)
            {
                sum += int.Parse(id17[i].ToString()) * int.Parse(id17Coefficient[i]);

            }
            int yushu = -1;
            Math.DivRem(sum, 11, out yushu);
            //if (remainderCode[yushu] != _id.Substring(17, 1).ToLower())
            //{
            //    return "错误";//校验码验证  

            //}
            return remainderCode[yushu].ToString();
        }
        private void getAge()
        {
            if (!_checkBoll)
            {
                _age = _checkStr;
            }
            else
            {
                int intYear = 0;                                    // 岁
                int intMonth = 0;                                    // 月
                int intDay = 0;                                    // 天
                DateTime dtNow = DateTime.Now;
                // 计算天数
                intDay = dtNow.Day - _birthday.Day;
                if (intDay < 0)
                {
                    dtNow = dtNow.AddMonths(-1);
                    intDay += DateTime.DaysInMonth(dtNow.Year, dtNow.Month);
                }

                // 计算月数
                intMonth = dtNow.Month - _birthday.Month;
                if (intMonth < 0)
                {
                    intMonth += 12;
                    dtNow = dtNow.AddYears(-1);
                }

                // 计算年数
                intYear = dtNow.Year - _birthday.Year;

                // 格式化年龄输出
                if (intYear >= 1)                                            // 年份输出
                {
                    _age = intYear.ToString() + "岁";
                }

                if (intMonth > 0 && intYear <= 5)                           // 五岁以下可以输出月数
                {
                    _age += intMonth.ToString() + "个月";
                }

                if (intDay >= 0 && intYear < 1)                              // 一岁以下可以输出天数
                {
                    if (intMonth == 0)
                    {
                        _age = intDay.ToString() + "天啦";
                    }
                    else
                    {
                        _age += "又" + intDay.ToString() + "天";
                    }
                }
            }
        }
        private void getSex()
        {
            if (!_checkBoll)
            {
                _sex = _checkStr;
            }
            else
            {
                int sexStr = 0;
                if (_id.Length == 15)
                {
                    sexStr = Convert.ToInt32(_id.Substring(14, 1));
                }
                else
                {
                    sexStr = Convert.ToInt32(_id.Substring(16, 1));
                }
                if (sexStr % 2 == 0)
                {
                    _sex = "女";
                }
                else
                {
                    _sex = "男";
                }
            }

        }

        private void getArea()
        {
            if (!_checkBoll)
            {
                _area = _checkStr;
            }
            else
            {
                int _num = int.Parse(_id.Substring(0, 6));
               
                _area =CheckNum(_num);
            }
        }
        /// <summary>
        /// 根据6位数编码查询地址
        /// </summary>
        /// <param name="number"></param>
        /// <returns></returns>
        private string CheckNum(int number)
        {
            System.Xml.XmlDocument xmldoc = new System.Xml.XmlDocument();
            xmldoc.LoadXml(Resource1.ResourceManager.GetObject("AreaNumber") as string);
            System.Xml.XmlNode nodecount = xmldoc.SelectSingleNode("//*[@id='" + number.ToString() + "']");
            if (nodecount != null)
            {
                if (number.ToString().Substring(2) == "0000")
                {
                    string c = nodecount.Attributes["name"].Value;
                    return c;
                }
                if (number.ToString().Substring(4) == "00")
                {
                    string a1 = nodecount.ParentNode.Attributes["name"].Value;
                    string a2 = nodecount.Attributes["name"].Value;
                    string a = a1 + a2;
                    return a;
                }
                string s1 = nodecount.ParentNode.ParentNode.Attributes["name"].Value;
                string s2 = nodecount.ParentNode.Attributes["name"].Value;
                string s3 = nodecount.Attributes["name"].Value;
                string s = s1 + s2 + s3;
                return s;
            }
            else
            {
                string idCity = number.ToString().Substring(0, 4) + "00";
                System.Xml.XmlNode nodecity = xmldoc.SelectSingleNode("//*[@id='" + idCity + "']");
                if (nodecity != null)
                {
                    string a1 = nodecity.ParentNode.Attributes["name"].Value;
                    string a2 = nodecity.Attributes["name"].Value;
                    string a = a1 + a2;
                    return a;
                }
                else
                {
                    string idProvinces = number.ToString().Substring(0, 2) + "0000";
                    System.Xml.XmlNode nodeProvinces = xmldoc.SelectSingleNode("//*[@id='" + idProvinces + "']");
                    if (nodeProvinces != null)
                    {
                        string c = nodeProvinces.Attributes["name"].Value;
                        return c;
                    }
                }
                return "未知地址！";
            }
        }
        #endregion

    }
}
