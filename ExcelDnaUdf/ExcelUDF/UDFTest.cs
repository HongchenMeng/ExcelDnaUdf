using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ExcelDnaUdf
{
    public partial class ExcelUDF
    {
        [ExcelFunction(Category = "自定义函数", IsMacroType = true, Description = "两列对应的单元格分别相加")]
        public static double[,] sumtest(double[,] s1, double[,] s2)
        {
            int rows = s1.GetLength(0);
            int columns = s1.GetLength(1);
            double[] s = new double[rows];
            for (int i = 0; i < rows; i++)
            {
                s[i] = s1[i, 0] + s2[i, 0];
            }
            ArrayResizer arrayResizer = new ArrayResizer();
            return arrayResizer.Resize(s);
        }

        // Just returns an array of the given size
        public static object[,] MakeArray(int rows, int columns)
        {
            object[,] result = new object[rows, columns];
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    result[i, j] = i + j;
                }
            }

            return result;
        }

        public static double[,] MakeArrayDoubles(int rows, int columns)
        {
            double[,] result = new double[rows, columns];
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    result[i, j] = i + (j / 1000.0);
                }
            }

            return result;
        }

        public static object MakeMixedArrayAndResize(int rows, int columns)
        {
            object[,] result = new object[rows, columns];
            for (int j = 0; j < columns; j++)
            {
                result[0, j] = "Col " + j;
            }
            for (int i = 1; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    result[i, j] = i + (j * 0.1);
                }
            }
            ArrayResizer arrayResizer = new ArrayResizer();
            return arrayResizer.Resize(result);
        }

        // Makes an array, but automatically resizes the result
        public static object MakeArrayAndResize(int rows, int columns, string unused, string unusedtoo)
        {
            object[,] result = MakeArray(rows, columns);
            ArrayResizer arrayResizer = new ArrayResizer();
            return arrayResizer.Resize(result);

            // Can also call Resize via Excel - so if the Resize add-in is not part of this code, it should still work
            // (though calling direct is better for large arrays - it prevents extra marshaling).
            // return XlCall.Excel(XlCall.xlUDF, "Resize", result);
        }

        public static double[,] MakeArrayAndResizeDoubles(int rows, int columns)
        {
            double[,] result = MakeArrayDoubles(rows, columns);
            ArrayResizer arrayResizer = new ArrayResizer();
            return arrayResizer.Resize(result);
        }
    }
}
