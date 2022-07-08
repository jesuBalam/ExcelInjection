using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelInjection
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Press any key to start.");
                Console.ReadKey();
                #region Excel Injection
                ExcelReader.ProcessExcel();
                Console.WriteLine("Process Done.");
                Console.ReadKey();
                #endregion
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
            }
        }
    }
}
