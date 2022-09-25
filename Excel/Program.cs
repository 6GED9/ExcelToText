using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelp.BO
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                using (BO.ExcelHelper helper = new BO.ExcelHelper())
                {
                    if (helper.Open(filePath: Path.Combine(Environment.CurrentDirectory, "nums.xlsx")))
                    {
                        helper.Get(column: "A", row: 1);
                        helper.Save();
                    }
                }
            }
            catch(Exception ex) { Console.WriteLine(ex.Message); }
            string text = null;
            using (FileStream stream = new FileStream("nums.txt", FileMode.Open))
            {
                for (int i = 1; i < ExcelHelper.range_values.Length; i++)
                {
                    text = ExcelHelper.range_values[i,1].ToString() + " \r\n";
                    byte[] array = System.Text.Encoding.Default.GetBytes(text);
                    stream.Write(array, 0, array.Length);
                }
            }
        }
    }
}
