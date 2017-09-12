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
        public static int RegexFindString(
             [ExcelArgument(Description = "源字符串")] string String_sentence,
            [ExcelArgument(Description = "查找目标")] string String_findStr,
            [ExcelArgument(Description = "第几次出现的位置")] int int_position)
        {
            int i = 0;
            try
            {
                if(int_position==0)
                {
                    int_position = 1;
                }
            foreach (Match match in Regex.Matches(String_sentence, String_findStr))
            {

                i++;
                if(i== int_position)
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
            return Regex.Replace(sentence, "[^0-9]", "", RegexOptions.IgnoreCase);
        }
        [ExcelFunction(Category = "文本处理", IsMacroType = true, Description = "提取字母")]
        public static string RegexReplaceAz(
    [ExcelArgument(Description = "源字符串")] string sentence)
        {
            return Regex.Replace(sentence, "[^a-zA-Z]", "", RegexOptions.IgnoreCase);
        }
        [ExcelFunction(Category = "文本处理", IsMacroType = true, Description = "提取中文")]
        public static string RegexReplaceChinese(
[ExcelArgument(Description = "源字符串")] string sentence)
        {
            return Regex.Replace(sentence, @"[^\u4e00-\u9fa5]+", "");
        }
        [ExcelFunction(Category = "文本处理", IsMacroType = true, Description = "替换文本")]
        public static string RegexReplaceOldNew(
[ExcelArgument(Description = "源字符串")] string sentence,
            [ExcelArgument(Description = "旧字符串")] string oldStr,
             [ExcelArgument(Description = "新字符串")] string newStr)
        {
            return Regex.Replace(sentence, oldStr, newStr);
        }

        [ExcelFunction(Category = "文本处理", IsMacroType = true, Description = "提取文本")]
        public static string RegexExtractString(
[ExcelArgument(Description = "源字符串")] string String_text,
    [ExcelArgument(Description = @"正则表达字符串：1为提取括号内的文本(?<=\()\S+(?=\))，  2为提取(?<=\()\S+(?=,)，   3为提取(?<=,)\S+(?=\)\.)，  4为提取%%号之后的文本 (?<=%%)\S+ ，5为提取逗号之前的文本 \S+(?=,)")] string String_Pattern)
        {
            switch(String_Pattern)
            {
                case "1":
                    String_Pattern = @"(?<=\()\S+(?=\))";
                    break;
                case "2":
                    String_Pattern = @"(?<=\()\S+(?=,)";
                    break;
                case "3":
                    String_Pattern = @"(?<=,)\S+(?=\)\.)";
                    break;
                case "4":
                    String_Pattern = @"(?<=%%)\S+";
                    break;
                case "5":
                    String_Pattern = @"\S+(?=,)";
                    break;
                default:
                    break;

            }
            MatchCollection mc = Regex.Matches(String_text, String_Pattern);
            string outStr = null;
            foreach (Match m in mc)
            {
                outStr = m.ToString();
                break;
            }

            return outStr;
        }
    }
}
