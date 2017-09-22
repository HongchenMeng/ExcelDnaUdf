using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


    //=GetFiles(A1,"\S\.xlsx$",TRUE,"L")
   public static class FileClass
    {
    /// <summary>
    /// 获取指定文件下所有【文件的全名称】，
    /// 可以获取子文件夹
    /// </summary>
    /// <param name="strFolder">指定目录</param>
    /// <param name="strRegesText">正则指定搜索条件,0为搜索xlsx文件，"\S\.xlsx$"表示 </param>
    /// <param name="blIsSearchSubfolder">是否搜索子文件夹,1搜索，0不搜索。默认不搜索</param>
    /// <returns>返回搜索到的文件全名数组</returns>
    public static string[] GetAllFiles( string strFolder, string strRegesText, bool blIsSearchSubfolder=false)
    {
        string[] files;
        if (strRegesText==null | strRegesText=="0")
        {
            strRegesText = @"\S\.xlsx$";
        }

        if(blIsSearchSubfolder)//搜索子文件夹
        {
            files = Directory.EnumerateFiles(strFolder, "*", SearchOption.AllDirectories).Where(s => isContainsText(s, strRegesText)).ToArray();
        }
        else
        {
            files = Directory.EnumerateFiles(strFolder).Where(s => isContainsText(s, strRegesText)).ToArray();
        }

        return files;
    }
    /// <summary>
    /// 
    /// </summary>
    /// <param name="strsFiles"></param>
    /// <param name="existExcelFileNames"></param>
    /// <param name="dic"></param>
    /// <returns></returns>
    public static bool FilesWithSameName(string[] strsFiles,out List<string> existExcelFileNames,out Dictionary<string, List<string>> dic)
    {
        existExcelFileNames = null;
        dic = null;
        foreach(string str in strsFiles)
        {
            string strP = Path.GetFileNameWithoutExtension(str);
            if (dic.ContainsKey(strP))//存在该key
            {
                dic[strP].Add(str);
            }
            else
            {
                dic.Add(strP, new List<string> { str });
            }
            existExcelFileNames.Add(strP);//不带扩展名的文件名称，如item
        }
        return false;
    }
    /// <summary>
    /// 获取指定文件下所有【文件夹】的名称，
    /// 可以获取子文件夹
    /// </summary>
    /// <param name="strFolder">指定目录</param>
    /// <param name="strRegesText">正则指定搜索条件 </param>
    /// <param name="blIsSearchSubfolder">是否搜索子文件夹,1搜索，0不搜索。默认不搜索</param>
    /// <returns>返回搜索到的文件全名数组</returns>
    public static string[] GetAllFolders(string strFolder, string strRegesText, bool blIsSearchSubfolder = false)
    {
        string[] files;
        if (string.IsNullOrEmpty(strRegesText.ToString().Trim()))
        {
            strRegesText = string.Empty;

        }

        if (blIsSearchSubfolder)//搜索子文件夹
        {
            files = Directory.EnumerateDirectories(strFolder, "*", SearchOption.AllDirectories).Where(s => isContainsText(s, strRegesText)).ToArray();
        }
        else
        {
            files = Directory.EnumerateDirectories(strFolder).Where(s => isContainsText(s, strRegesText)).ToArray();
        }

        return files;
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

