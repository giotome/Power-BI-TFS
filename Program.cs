using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using System.Configuration;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;

namespace RefreshExcel1
{
    class Program
    {
        static void Openfile()
        {
            string mySheet = ConfigurationSettings.AppSettings["EXCEL_PATH"];
            var excelApp = new Excel.Application();
            excelApp.Visible = true;

            Excel.Workbooks books = excelApp.Workbooks;

            Excel.Workbook sheet = books.Open(mySheet);

            foreach (CommandBar bar in sheet.CommandBars)
            {
                if ("Team".Equals(bar.Name))
                {
                    foreach (CommandBarControl control in bar.Controls)
                    {
                        if ("IDC_REFRESH".Equals(control.Tag.ToUpper()))
                            control.Execute();
                    }
                }
            }


            sheet.RefreshAll();
            sheet.Save();
            excelApp.Quit();

        }

        static void Main(string[] args)
        {
            Openfile();
        }
    }
}
