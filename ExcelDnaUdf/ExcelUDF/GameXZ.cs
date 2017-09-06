using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace ExcelDnaUdf
{
    public partial class ExcelUDF
    {
        [ExcelFunction(Category = "自定义函数", IsMacroType = true, Description = "缘分配置1")]
        public static object VLOOKUPMERGE(double[,] s1, double[,] s2,double[,] s3)
        {
            int rows2 = s2.GetLength(0);
            int columns2 = s2.GetLength(1);

            int rows3 = s3.GetLength(0);
            int columns3 = s3.GetLength(1);

            if(rows2 !=rows3 || columns2 !=columns3)
                return ExcelError.ExcelErrorValue;

            if(rows2 >1 && columns2 >1)
                return ExcelError.ExcelErrorValue;


            Dictionary<string, string> dic = new Dictionary<string, string>();

            if (rows2 >1 && columns2 ==1)
            {
                for (int i = 0; i < rows2; i++)
                {
                    string srid = s2[i,0].ToString();
                    string srv = s3[i, 0].ToString();
                    if(dic.ContainsKey(srid))
                    {
                        dic[srid] = dic[srid] + "," + srv;
                    }
                    else
                    {
                        dic.Add(srid, srv);
                    }
                }
            }


            int rows = s1.GetLength(0);
            int columns = s1.GetLength(1);
            object[] ResultObject = new object[rows];
            for (int i = 0; i < rows; i++)
            {
                ResultObject[i] = dic[s1[i,0].ToString()];
            }
            ArrayResizer arrayResizer = new ArrayResizer();
            return arrayResizer.Resize(ResultObject);
        }
    }
}
