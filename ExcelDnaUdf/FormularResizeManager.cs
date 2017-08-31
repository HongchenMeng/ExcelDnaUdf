using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelDnaUdf
{
    class FormularResizeManager
    {
        public Range SrcRangeItem { get; set; }
        public bool IsHAlign { get; set; }

        internal void RizeFormularArrayRange()
        {
            object arrData = GetDataFromUDF();
            if (arrData is Array)
            {
                FillDownFormularFromResultData(arrData);
            }
        }



        private object GetDataFromUDF()
        {
            try
            {
                string formularString = this.SrcRangeItem.Formula;
                var matchs = Regex.Matches(formularString, "\".+?\"");
                int i = 0;
                foreach (Match item in matchs)
                {
                    i++;
                    var itemValue = item.Value;
                    formularString = formularString.Remove(formularString.IndexOf(itemValue), itemValue.Length).Insert(formularString.IndexOf(itemValue), "#" + i.ToString());
                }
                string formular = formularString.Split(new char[] { '(' })[0].TrimStart(new char[] { '=' });
                string formularPara = formularString.Split(new char[] { '(' })[1].Replace(")", "");
                string[] paras = formularPara.Split(new char[] { ',' });
                //把之前的字符串重新换回来
                int j = 0;
                foreach (var item in paras.Where(s => s.ToString().Contains("#")))
                {
                    paras[Array.IndexOf(paras, item)] = matchs[j].Value;
                    j++;
                }
                //把参数装入LIST中，其中是Range对象的参数要处理成Range对象装入，是字符串的要去掉双引号。
                List<object> listparas = new List<object>();
                //加入函数名
                listparas.Add(formular);
                //加入函数参数
                listparas.AddRange(CreateParasListOfFormular(paras));
                var result = Common.RunMacro(Common.xlApp, listparas.ToArray());
                return result;
                //return Common.RunMacro(Common.xlApp, listparas.ToArray());
            }
            catch (Exception)
            {
                throw;
            }

        }

        private void FillDownFormularFromResultData(object arrData)
        {
            //只有返回是数组时才处理，否则只是普通公式不作处理
            if (arrData is Array)
            {
                int lBound0 = ((Array)arrData).GetLowerBound(0);
                int uBound0 = ((Array)arrData).GetUpperBound(0);
                if (this.IsHAlign)
                {
                    this.SrcRangeItem.Resize[1, uBound0 - lBound0 + 1].FormulaArray = this.SrcRangeItem.Formula;
                }
                else
                {
                    this.SrcRangeItem.Resize[uBound0 - lBound0 + 1, 1].FormulaArray = this.SrcRangeItem.Formula;
                }

            }
        }

        private List<object> CreateParasListOfFormular(string[] paras)
        {
            List<object> listparas = new List<object>();
            for (int k = 0; k < paras.Length; k++)
            {
                //非字符串时，没有双引号引住内容
                if (paras[k].Trim(new char[] { '"' }) == paras[k])
                {
                    //传入的是EXCEL的布尔值
                    if (paras[k].Trim().Equals("TRUE", StringComparison.CurrentCultureIgnoreCase) || paras[k].Trim().Equals("FALSE", StringComparison.CurrentCultureIgnoreCase))
                    {

                    }
                    //当为单元格时，如A1或A5：B6
                    else if (Regex.IsMatch(paras[k].Trim(), @"[a-zA-Z]+\d+$|[a-zA-Z]+\d+:[a-zA-Z]+\d+$"))
                    {
                        listparas.Add(Common.xlApp.Range[paras[k].Trim()]);
                    }
                    //当为名称引用时，报错，因为Run方法不支持名称操作，.net正则\w支持中文。
                    else if (Regex.IsMatch(paras[k], @"\w+"))
                    {
                        throw new Exception("传入参数出错，公式扩展时不允许用定义名称来传入");
                    }
                    //数字直接加
                    else
                    {
                        listparas.Add(paras[k].Trim());
                    }

                }
                //字符串时
                else
                {
                    string paraString = paras[k].Trim(new char[] { '"', '\0' });
                    if (paraString.Equals("H", StringComparison.CurrentCultureIgnoreCase))
                    {
                        listparas.Add("L");
                        //做一下标识，让后续的方法以H的方式来处理。
                        this.IsHAlign = true;
                    }
                    else
                    {
                        listparas.Add(paraString);
                    }
                }
            }

            return listparas;
        }

        /// <summary>
        /// 获取要进行区域扩展的公式单元格区域装入list中供后续调用
        /// </summary>
        /// <param name="selectRange"></param>
        /// <returns></returns>
        public static List<Range> GetRangeOfFormular(Range selectRange)
        {
            List<Range> firstRangeOfFormularArrays = new List<Range>();
            foreach (Range item in selectRange.Cells)
            {
                //当单元格为数组公式一部分时
                if (item.HasArray)
                {
                    //只保留第1个公式的内容，其他公式内容删除
                    string formularString = item.Formula;
                    //只处理数组公式，如果当前引用的公式已经包含数组公式，就只处理第1个单元格。
                    Range firtRangeOfFormularArray = item.CurrentArray[1];
                    firstRangeOfFormularArrays.Add(firtRangeOfFormularArray);
                    //删除数组公式并清空内容，再重新赋值会首个单元格为数组公式的内容
                    item.CurrentArray.ClearContents();
                    //重新填充第1个单元格公式
                    firtRangeOfFormularArray.Formula = formularString;
                }
                //只处理有公式的部分，前面已经清空了数组公式引用的非首个单元格，此处会忽略他们
                //当公式返回多个值，但用户没有用数组公式输入时，HasArray属性为假，HasFormula为真，这部分也要包含进去。
                else if (item.HasFormula)
                {
                    firstRangeOfFormularArrays.Add(item);
                }
            }
            return firstRangeOfFormularArrays;
        }
    }
}
