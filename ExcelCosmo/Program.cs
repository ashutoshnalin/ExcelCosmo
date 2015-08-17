using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel;

namespace ExcelCosmo
{
    class Program
    {
        static void Main(string[] args)
        {
            // UnComment it when used to run .exe file directly.
            //FileInfo fileInfo = new FileInfo(args[0]);
            try
            {
                // For testing below two lines
                // D:\ClientProjects\Elance\ExcelCosmo\ExcelCosmo\bin\Debug
                FileInfo fileInfo = new FileInfo("D:\\ClientProjects\\Elance\\ExcelCosmo\\ExcelCosmo\\bin\\Debug\\Cosmetologist.xlsx");
                if (fileInfo != null)
                {
                    string fileName = Path.GetFileNameWithoutExtension(fileInfo.Name);
                    ExcelReaderCreator.CreateComosExcel(fileInfo.Name);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
