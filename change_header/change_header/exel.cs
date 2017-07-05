using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;

namespace change_header
{
    interface IReadChangeFile
    {
        void changeDataOfFile(Dictionary<string,string> dict);
        void openAnotherFile();
        void closeFile();
    }

    public class exel : IReadChangeFile
    {
        public Excel.Application xlApp;
        public Excel.Workbook xlWorkbook;
        public Excel._Worksheet xlWorksheet;
        public Excel.Range xlRange;
        public exel(string name)
        {
            this.xlApp = new Excel.Application();
            this.xlWorkbook = this.xlApp.Workbooks.Open(name);
            this.xlWorksheet = this.xlWorkbook.Sheets[1];
            xlWorksheet.Name = ConfigurationSettings.AppSettings["newworksheet"];
            this.xlRange = this.xlWorksheet.UsedRange;
        }
        public void changeDataOfFile(Dictionary<string,string> dict)
        {
            for (int i = 1; i <= xlRange.Columns.Count; i++)
            {
                xlRange[1, i].value2 = dict[xlRange[1, i].value2];
            }
        }
        public void openAnotherFile()
        {
            xlWorksheet = xlWorkbook.Sheets.Add();
            xlWorksheet.Name = ConfigurationSettings.AppSettings["oldworksheet"]; 
            for (int i = 1; i <= xlRange.Rows.Count; i++)
            {
                for (int j = 1; j <= xlRange.Columns.Count; j++)
                {
                    xlWorksheet.Cells[j, i].value2 = xlRange[j, i].value2;

                    // Console.WriteLine("i={0},j={1}", i, j);
                }
            }
        }
        public void closeFile()
        {
            xlApp.ActiveWorkbook.Save();
            xlWorkbook.Close(false, Type.Missing, Type.Missing);
            xlApp.Quit();
            GC.Collect();
        }
        
    }
}
