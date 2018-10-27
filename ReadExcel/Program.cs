using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            var r = new Read(@"C:\\Users\\alroc\\Desktop\\table_data.xlsx");
            //r.ReadRows();
            r.CreateJson();
        }
    }
}
