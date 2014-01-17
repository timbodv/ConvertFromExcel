using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Timbodv.ConvertFromExcel
{
    public class Program
    {
        [STAThread]
        public static void Main(string[] args)
        {
            var result = ConvertExcelHelper.ConvertExcelToDataset(@"C:\Users\Tim\Documents\sample.xls", false);
            Console.WriteLine(result);
        }
    }
}
