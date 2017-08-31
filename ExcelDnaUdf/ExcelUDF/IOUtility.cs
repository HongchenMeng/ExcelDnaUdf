using System;
using System.IO;
using System.Text;
using System.Linq;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using System.Collections.Generic;
using System.Threading.Tasks;


namespace ExcelDnaUdf
{
    public partial class ExcelUDF
    {

        [ExcelFunction(Category = "文件处理", IsMacroType = true, Description = "获取指定目录下的子文件夹,srcFolder为传入的顶层目录")]
        public static object GetSubFolders(
            [ExcelArgument(Description = "传入的顶层目录，最终返回的结果将是此目录下的文件夹或子文件夹")] string FileDirectory,
            [ExcelArgument(Description = "查找的文件夹中是否需要包含指定字符串，不传参数默认为返回所有文件夹，可传入复杂的正则表达式匹配。")] string optContainsText,
            [ExcelArgument(Description = "是否查找所有子文件夹，TRUE搜索子文件夹，fasle为否，默认为否")] bool optIsSearchAllDirectory=false,
            [ExcelArgument(Description = "返回的结果是按列排列还是按行排列，true按行，false按列，默认按列")] bool isHorizontal=false)
        {

            string[] subfolders;
            if (Common.IsMissOrEmpty(optContainsText))
            {
                optContainsText = string.Empty;
            }
            //当isSearchAllDirectory为空或false，默认为只搜索顶层文件夹
            if (optIsSearchAllDirectory== false)
            {
                subfolders = Directory.EnumerateDirectories(FileDirectory).Where(s => isContainsText(s, optContainsText)).ToArray();
            }
            else
            {

                subfolders = Directory.EnumerateDirectories(FileDirectory, "*", SearchOption.AllDirectories).Where(s => isContainsText(s, optContainsText)).ToArray();
            }
            ArrayResizer arrayResizer = new ArrayResizer();
            return arrayResizer.Resize(subfolders, isHorizontal);

        }


        [ExcelFunction(Category = "文件处理", IsMacroType = true, Description = "获取指定目录下文件的文件全名")]
        public static object GetFiles(
                [ExcelArgument(Description = "传入的顶层目录，最终返回的结果将是此目录下的文件夹或子文件夹下的全路径文件名")] string srcFolder,
                [ExcelArgument(Description = "查找的全路径文件名中是否需要包含指定字符串，不传参数默认为返回所有文件夹，可传入复杂的正则表达式匹配。")] string containsText,
                [ExcelArgument(Description = "是否查找顶层目录下的文件夹的所有子文件夹，TRUE和非0的字符或数字为搜索子文件夹，其他为否，不传参数时默认为否")] object isSearchAllDirectory,
                [ExcelArgument(Description = "返回的结果是按按列排列还是按行排列，传入L按列排列，传入H按行排列，不传参数或传入非L或H则默认按列排列")] string optAlignHorL)
        {
            string[] files;
            if (Common.IsMissOrEmpty(containsText))
            {
                containsText = string.Empty;
            }
            //当isSearchAllDirectory为空或false，默认为只搜索顶层文件夹
            if (Common.IsMissOrEmpty(isSearchAllDirectory) || Common.TransBoolPara(isSearchAllDirectory) == false)
            {
                files = Directory.EnumerateFiles(srcFolder).Where(s => isContainsText(s, containsText)).ToArray();
            }
            else
            {

                files = Directory.EnumerateFiles(srcFolder, "*", SearchOption.AllDirectories).Where(s => isContainsText(s, containsText)).ToArray();
            }
            ArrayResizer arrayResizer = new ArrayResizer();
            return arrayResizer.Resize(Common.ReturnDataArray(files, optAlignHorL));
        }

        private static bool isContainsText(string s, string containstext)
        {
            if (string.IsNullOrEmpty(containstext))
            {
                return true;
            }
            else
            {
                return System.Text.RegularExpressions.Regex.IsMatch(s, containstext);
            }

        }
    }
}
