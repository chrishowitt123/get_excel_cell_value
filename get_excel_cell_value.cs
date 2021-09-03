
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelApp = Microsoft.Office.Interop.Excel;



namespace RenameExcelFileFromTab
{
    //dont forget -- using Microsoft.Office.Interop.Excel;
    class Program
    {
        static void Main(string[] args)
        {
            var path = @"C:\Users\CHowitt01\Downloads\Wait List.xlsx";
            using (var package = new ExcelPackage(new FileInfo(path)))

            {
                var firstSheet = package.Workbook.Worksheets["WL Done"];
                Console.WriteLine($"Cell A2 Value   : {firstSheet.Cells["G2"].Text}");

            }

        }
    }
}
