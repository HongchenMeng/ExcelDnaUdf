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
            subfolders = FileClass.GetAllFolders(FileDirectory, optContainsText, optIsSearchAllDirectory);
            //if (Common.IsMissOrEmpty(optContainsText))
            //{
            //    optContainsText = string.Empty;
            //}
            ////当isSearchAllDirectory为空或false，默认为只搜索顶层文件夹
            //if (optIsSearchAllDirectory== false)
            //{
            //    subfolders = Directory.EnumerateDirectories(FileDirectory).Where(s => isContainsText(s, optContainsText)).ToArray();
            //}
            //else
            //{

            //    subfolders = Directory.EnumerateDirectories(FileDirectory, "*", SearchOption.AllDirectories).Where(s => isContainsText(s, optContainsText)).ToArray();
            //}
            ArrayResizer arrayResizer = new ArrayResizer();
            return arrayResizer.Resize(subfolders, isHorizontal);

        }


        [ExcelFunction(Category = "文件处理", IsMacroType = true, Description = "获取指定文件下所有文件的全名称，可以获取子文件夹")]
        public static object GetFiles(
                [ExcelArgument(Description = "指定目录")] string strFolder,
                [ExcelArgument(Description = @"正则指定搜索条件,0为搜索xlsx文件，\S\.xlsx$ 表示")] string strRegesText,
                [ExcelArgument(Description = "是否搜索子文件夹,1搜索，0不搜索。默认不搜索")] bool blIsSearchSubfolder,
                [ExcelArgument(Description = "返回的结果是按【行】排列，true按行，false按列，默认按列")] bool blIsHorizontal)
        {
            string[] files;
            files = FileClass.GetAllFiles(strFolder, strRegesText, blIsSearchSubfolder);
            //if (Common.IsMissOrEmpty(containsText))
            //{
            //    containsText = string.Empty;
            //}
            ////当isSearchAllDirectory为空或false，默认为只搜索顶层文件夹
            //if (Common.IsMissOrEmpty(isSearchAllDirectory) || Common.TransBoolPara(isSearchAllDirectory) == false)
            //{
            //    files = Directory.EnumerateFiles(srcFolder).Where(s => isContainsText(s, containsText)).ToArray();
            //}
            //else
            //{

            //    files = Directory.EnumerateFiles(srcFolder, "*", SearchOption.AllDirectories).Where(s => isContainsText(s, containsText)).ToArray();
            //}
            ArrayResizer arrayResizer = new ArrayResizer();
            //return arrayResizer.Resize(Common.ReturnDataArray(files, optAlignHorL));
            return arrayResizer.Resize(files,blIsHorizontal);
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
