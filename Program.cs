using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace RefreshExcel1
{
    class Program
    {
        static void Openfile()
        {
            string mySheet = @"C:\Users\g.pettenazzi.tome\OneDrive - Avanade\Giovanna\EINSTEIN\Power BI e TFS\DSS - Escopo do Projeto.xlsx";
            var excelApp = new Excel.Application();
            excelApp.Visible = true;

            Excel.Workbooks books = excelApp.Workbooks;

            Excel.Workbook sheet = books.Open(mySheet);

            sheet.RefreshAll();
            sheet.Save();

        }
        static void Main(string[] args)
        {
            Openfile();
        }
    }
}
