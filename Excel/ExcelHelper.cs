using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelHelp.BO
{
    class ExcelHelper:  IDisposable
    {
        private Application _excel;
        private Workbook _workbook;
        public object num;
        Excel.Range range;
        public static object[,] range_values;
        private string _filePath;

        public ExcelHelper()
        {
            _excel = new Excel.Application();
        }


        internal bool Open(string filePath)
        {
            try
            {
               if (File.Exists(filePath))
                {
                    _workbook = _excel.Workbooks.Open(filePath);
                    _filePath = filePath;
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return true;
        }

        internal void Save()
        {
            if(!string.IsNullOrEmpty(_filePath))
            {
                _workbook.SaveAs(_filePath);
                _filePath = null;
            }
            else
            {
                _workbook.Save();
            }
        }

        internal bool Get(string column, int row)
        {
            try
            {
                range = (Excel.Range)_excel.Columns[column, Type.Missing];
                Excel.Range last_cell = range.get_End(Excel.XlDirection.xlDown);
                Excel.Range first_cell = (Excel.Range)_excel.Cells[row, column];
                Excel.Range value_range = (Excel.Range)_excel.get_Range(first_cell, last_cell);
                range_values = (object[,])value_range.Value2;
                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return false;
        }
        public void Dispose()
        {
            try
            {
                _workbook.Close();
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
    }
