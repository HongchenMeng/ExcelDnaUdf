using System;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;


namespace ExcelDnaUdf
{
    public partial class ExcelUDF
    {
        [ExcelFunction(Category = "文本处理", IsMacroType = true, Description = "查找字符串的位置")]
        public static int RegexMatches(
             [ExcelArgument(Description = "源字符串")] string sentence,
            [ExcelArgument(Description = "查找目标")] string findStr,
            [ExcelArgument(Description = "第几次出现的位置")] int position)
        {
            int i = 0;
            try
            {
                if(position==0)
                {
                    position = 1;
                }
            foreach (Match match in Regex.Matches(sentence, findStr))
            {

                i++;
                if(i== position)
                {
                    return match.Index;
                }
            }

            }
            catch
            {
                return -1;
            }
            return -1;

        }

        [ExcelFunction(Category = "文本处理", IsMacroType = true, Description = "提取数字")]
        public static string RegexReplaceNum(
            [ExcelArgument(Description = "源字符串")] string sentence)
        {
            return Regex.Replace(sentence, "[0-9]", "", RegexOptions.IgnoreCase);
        }
        [ExcelFunction(Category = "文本处理", IsMacroType = true, Description = "提取字母")]
        public static string RegexReplaceAz(
    [ExcelArgument(Description = "源字符串")] string sentence)
        {
            return Regex.Replace(sentence, "[a-zA-Z]", "", RegexOptions.IgnoreCase);
        }
        [ExcelFunction(Category = "文本处理", IsMacroType = true, Description = "提取中文")]
        public static string RegexReplaceChinese(
[ExcelArgument(Description = "源字符串")] string sentence)
        {
            return Regex.Replace(sentence, @"[\u4e00-\u9fa5]+", "");
        }
        [ExcelFunction(Category = "文本处理", IsMacroType = true, Description = "替换文本")]
        public static string RegexReplaceOldNew(
[ExcelArgument(Description = "源字符串")] string sentence,
            [ExcelArgument(Description = "旧字符串")] string oldStr,
             [ExcelArgument(Description = "新字符串")] string newStr)
        {
            return Regex.Replace(sentence, oldStr, newStr);
        }
    }
}
