using OfficeOpenXml;
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
            if (parts.Length != 3)
            {
                Console.WriteLine("Command not valid, Combine requires the name of both reports (including name and file extension) and the type of report for formatting.");
                return;
            }
            var file1 = parts[1];
            var file2 = parts[2];
            //var type = parts[3].ToLower();

            var files = new string[] { file1, file2 };

            var resultFile = @"C:\Users\elusk\Desktop\result.xlsx";

            ExcelPackage masterPackage = new ExcelPackage(new FileInfo(resultFile));
            foreach (var file in files)
            {
                ExcelPackage pckg = new ExcelPackage(new FileInfo(file));

                foreach (var sheet in pckg.Workbook.Worksheets)
                {
                    //check name of worksheet, in case that worksheet with same name already exist exception will be thrown by EPPlus

                    string workSheetName = sheet.Name;
                    foreach (var masterSheet in masterPackage.Workbook.Worksheets)
                    {
                        if (sheet.Name == masterSheet.Name)
                        {
                            workSheetName = string.Format("{0}_{1}", workSheetName, DateTime.Now.ToString("yyyyMMddhhssmmm"));
                        }
                    }

                    //add new sheet
                    masterPackage.Workbook.Worksheets.Add(workSheetName, sheet);
                }
            }

            masterPackage.SaveAs(new FileInfo(resultFile));
            Console.WriteLine("File created");
        }

    }
}
