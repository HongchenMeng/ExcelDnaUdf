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
        //--input=输入
        //--pattern=匹配规则
        //--matchNum=确定第几个匹配返回值，索引号从0开始，第1个匹配，传入0
        //--groupNum=确定第几组匹配，索引号从1开始，0为返回上层的match内容。
        //--isCompiled=是否编译，是为1，否为0，暂时没有测试过哪个快在数据量大时，文档好像说数据量大用编译比较好
        //--isECMAScript，用来指定\w是否匹配一些特殊编码之类的例如中文，当false时会匹配中文,指定为true时，可能和其他的指定有些冲突。
        //--RegexOptions.ECMAScript 选项只能与 RegexOptions.IgnoreCase 和 RegexOptions.Multiline 选项结合使用。在正则表达式中使用其他选项会导致 ArgumentOutOfRangeException。
        //--isRightToLeft，从右往左匹配。
        //--returnNum，反回split数组中的第几个元素，索引从0开始

        [ExcelFunction(Category = "文本正则", IsMacroType = true, Description = "正则匹配组，Pattern里传入（）来分组")]
        public static string RegexMatchGroup(
           [ExcelArgument(Description = "输入的字符串")] string input,
           [ExcelArgument(Description = "匹配规则")] string pattern,
           [ExcelArgument(Description = "确定第几个匹配返回值，索引号从0开始，第1个匹配，传入0，默认为0")] int matchNum = 0,
           [ExcelArgument(Description = "确定第几组匹配，索引号从1开始，0为返回上层的match内容，默认为1")] int groupNum = 1,
           [ExcelArgument(Description = "是否编译，是为1，否为0，暂时没有测试过哪个快在数据量大时，文档好像说数据量大用编译比较好，默认为false")] bool isCompiled = false,
           [ExcelArgument(Description = @"用来指定\w是否匹配一些特殊编码之类的例如中文，当false时会匹配中文,指定为true时，可能和其他的指定有些冲突，默认为false")] bool isECMAScript = false,
           [ExcelArgument(Description = "从右往左匹配，默认为false")] bool isRightToLeft = false)
        {
            try
            {
                RegexOptions options = GetRegexOptions(isCompiled, isECMAScript, isRightToLeft);
                MatchCollection matches = Regex.Matches(input, pattern, options);

                if (matchNum <= matches.Count - 1)
                {
                    Match match = matches[matchNum];
                    if (groupNum == 0)
                    {
                        return match.Value;
                    }
                    else
                    {
                        if (groupNum < match.Groups.Count)
                        {
                            return match.Groups[groupNum].Value;
                        }
                        else
                        {
                            return "";
                        }
                    }
                }
                else
                {
                    return "";
                }

            }
            catch (ArgumentOutOfRangeException)
            {
                return "Options错误";
            }
            catch (ArgumentNullException)
            {
                return "";
            }

            catch (ArgumentException)
            {
                return "Pattern错误";
            }
            catch (Exception)
            {
                return "";
            }
        }

        [ExcelFunction(Category = "文本正则", IsMacroType = true, Description = "正则匹配，不含Group组匹配")]
        public static string RegexMatch(
           [ExcelArgument(Description = "输入的字符串")] string input,
           [ExcelArgument(Description = "匹配规则")] string pattern,
           [ExcelArgument(Description = "确定第几个匹配返回值，索引号从0开始，第1个匹配，传入0，默认为0")] int matchNum = 0,
           [ExcelArgument(Description = "是否编译，是为1，否为0，暂时没有测试过哪个快在数据量大时，文档好像说数据量大用编译比较好，默认为false")] bool isCompiled = false,
           [ExcelArgument(Description = @"用来指定\w是否匹配一些特殊编码之类的例如中文，当false时会匹配中文,指定为true时，可能和其他的指定有些冲突，默认为false")] bool isECMAScript = false,
           [ExcelArgument(Description = "从右往左匹配，默认为false")] bool isRightToLeft = false)
        {
            int groupNum = 0;
            return RegexMatchGroup(input, pattern, matchNum, groupNum, isCompiled, isECMAScript, isRightToLeft);
        }

        [ExcelFunction(Category = "文本正则", IsMacroType = true, Description = "正则替换")]
        public static string RegexReplace(
           [ExcelArgument(Description = "输入的字符串")] string input,
           [ExcelArgument(Description = "匹配规则")] string pattern,
           [ExcelArgument(Description = "匹配到的文件替换的字符串，默认为替换为空")] string replacement = "",
           [ExcelArgument(Description = "是否编译，是为1，否为0，暂时没有测试过哪个快在数据量大时，文档好像说数据量大用编译比较好，默认为false")] bool isCompiled = false,
           [ExcelArgument(Description = @"用来指定\w是否匹配一些特殊编码之类的例如中文，当false时会匹配中文,指定为true时，可能和其他的指定有些冲突，默认为false")] bool isECMAScript = false,
           [ExcelArgument(Description = "从右往左匹配，默认为false")] bool isRightToLeft = false)
        {
            try
            {
                RegexOptions options = GetRegexOptions(isCompiled, isECMAScript, isRightToLeft);
                return Regex.Replace(input, pattern, replacement, options);
            }
            catch (ArgumentOutOfRangeException)
            {
                return "Options错误";
            }
            catch (ArgumentNullException)
            {
                return "";
            }

            catch (ArgumentException)
            {
                return "Pattern错误";
            }
            catch (Exception)
            {
                return input;
            }
        }

        /// <summary>
        /// 文本分割
        /// </summary>
        /// <param name="input"></param>
        /// <param name="pattern"></param>
        /// <param name="returnNum">索引从0开始</param>
        /// <param name="isCompiled"></param>
        /// <param name="isECMAScript"></param>
        /// <param name="isRightToLeft"></param>
        /// <returns></returns>
        [ExcelFunction(Category = "文本正则", IsMacroType = true, Description = "正则分割")]
        public static string RegexSplit(
           [ExcelArgument(Description = "输入的字符串")] string input,
           [ExcelArgument(Description = "匹配规则")] string pattern,
           [ExcelArgument(Description = "分割后返回第几个项目，索引号从0开始，第1个匹配，传入0，，默认为0")] int returnNum = 0,
           [ExcelArgument(Description = "是否编译，是为1，否为0，暂时没有测试过哪个快在数据量大时，文档好像说数据量大用编译比较好，默认为false")] bool isCompiled = false,
           [ExcelArgument(Description = @"用来指定\w是否匹配一些特殊编码之类的例如中文，当false时会匹配中文,指定为true时，可能和其他的指定有些冲突，默认为false")] bool isECMAScript = false,
           [ExcelArgument(Description = "从右往左匹配，默认为false")] bool isRightToLeft = false)
        {
            try
            {
                RegexOptions options = GetRegexOptions(isCompiled, isECMAScript, isRightToLeft);
                string[] splitResult = Regex.Split(input, pattern, options);
                if (returnNum <= splitResult.Length - 1)
                {
                    return splitResult[returnNum];
                }
                else
                {
                    return "";
                }

            }
            catch (ArgumentOutOfRangeException)
            {
                return "Options错误";
            }
            catch (ArgumentNullException)
            {
                return "";
            }

            catch (ArgumentException)
            {
                return "Pattern错误";
            }
            catch (Exception)
            {
                return input;
            }
        }

        [ExcelFunction(Category = "文本正则", IsMacroType = true, Description = "正则匹配判断")]
        public static bool RegexIsMatch(
           [ExcelArgument(Description = "输入的字符串")] string input,
           [ExcelArgument(Description = "匹配规则")] string pattern,
           [ExcelArgument(Description = "是否编译，是为1，否为0，暂时没有测试过哪个快在数据量大时，文档好像说数据量大用编译比较好，默认为false")] bool isCompiled = false,
           [ExcelArgument(Description = @"用来指定\w是否匹配一些特殊编码之类的例如中文，当false时会匹配中文,指定为true时，可能和其他的指定有些冲突，默认为false")] bool isECMAScript = false,
           [ExcelArgument(Description = "从右往左匹配，默认为false")] bool isRightToLeft = false)
        {
            try
            {
                RegexOptions options = GetRegexOptions(isCompiled, isECMAScript, isRightToLeft);
                return Regex.IsMatch(input, pattern, options);
            }

            catch (Exception)
            {
                return false;
            }
        }

        private static RegexOptions GetRegexOptions(bool isCompiled, bool isECMAScript, bool isRightToLeft)
        {
            List<RegexOptions> listOptions = new List<RegexOptions>();
            if (isCompiled == true)
            {
                listOptions.Add(RegexOptions.Compiled);
            }
            if (isRightToLeft == true)
            {
                listOptions.Add(RegexOptions.RightToLeft);
            }
            if (isECMAScript == true)
            {
                listOptions.Add(RegexOptions.ECMAScript);
            }

            RegexOptions options = new RegexOptions();
            foreach (var item in listOptions)
            {
                if (options == 0)
                {
                    options = item;
                }
                else
                {
                    options = options | item;
                }
            }

            return options;
        }


        [ExcelFunction(Category = "文本正则", IsMacroType = true, Description = "判断字符是否在指定的字符串集合内，如果存在返回true，否则返回false，示例：sourceStrings=AB,CD,E,lookupValue=CD,strSplit=',',CD在｛AB，CD，E｝的集合中，返回true")]
        public static bool IsTextContainsWithSplit(
             [ExcelArgument(Description = "查找字符串集合")] string sourceStrings,
             [ExcelArgument(Description = "查找条件")] string lookupValue,
             [ExcelArgument(Description = "查找字符串集合内用于分割的字符，注意中英文符号要与查找字符串集合一致,若传入多个分隔符，使用|隔开。")] string strSplit
    )
        {
            string[] strsplits;
            //当传入的strsplit是以|结尾或开头的，就当作一个字符串处理，不进行strsplit的分隔
            if (strSplit.Trim(new char[] { '|' }).Length != strSplit.Length || strSplit == "|")
            {
                strsplits = new string[] { strSplit.Trim() };
            }
            else
            {
                strsplits = strSplit.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray();
            }
            return sourceStrings.Split(strsplits, StringSplitOptions.RemoveEmptyEntries).Contains(lookupValue);

        }
    }
}
