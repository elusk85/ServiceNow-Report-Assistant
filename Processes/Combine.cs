using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System;
using System.Collections.Generic;
using System.IO;

namespace ServiceNow_Report_Assistant.Processes
{
    class Combine
    {
        public void CombineFiles(string command)
        {
            var parts = command.Split(',');
            if (parts.Length != 3) //needs to be changed to 4 when ready for the type argument
            {
                Console.WriteLine("Command not valid, Combine requires the name of both reports (including name and file extension) and the type of report for formatting.");
                return;
            }
            var file1 = parts[1];
            var file2 = parts[2];
            //var type = parts[3].ToLower(); //not needed yet

            var files = new string[] { file1, file2 };

            var resultFile = @"C:\Users\elusk\Desktop\Combined_Report_" + DateTime.Now.ToString("yyyyMMddhhssmmm") + ".xlsx";

            ExcelPackage masterPackage = new ExcelPackage(new FileInfo(resultFile));
            var ws = masterPackage.Workbook.Worksheets.Add("Sheet1"); //possibly not set properly
            ExcelRangeBase rangeBase = ws.Cells[1, 1];

            foreach (var file in files)
            {
                ExcelPackage pckg = new ExcelPackage(new FileInfo(file));

                foreach (var sheet in pckg.Workbook.Worksheets) //-this section puts the worksheets into a single .xlsx file
                {
                    //check name of worksheet, in case that worksheet with same name already exist exception will be thrown by EPPlus
                    string workSheetName = sheet.Name;
                    foreach (var masterSheet in masterPackage.Workbook.Worksheets)
                    {
                            int rowcount = sheet.Dimension.End.Row;
                            int colcount = sheet.Dimension.End.Column;

                            sheet.Cells[2, 1, rowcount, colcount].Copy(rangeBase);//possibly not set properly
                    }

                    //add new sheet if possible
                    //try
                    //{
                    //    masterPackage.Workbook.Worksheets.Add(workSheetName, sheet);
                    //}
                    //catch (InvalidOperationException e)
                    //{
                    //    Console.WriteLine("Exception: " + e );
                    //}

                }
            }



            masterPackage.SaveAs(new FileInfo(resultFile));
            Console.WriteLine("File created");
        }

    }
}
