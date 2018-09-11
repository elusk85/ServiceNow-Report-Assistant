﻿using OfficeOpenXml;
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
            foreach (var file in files)
            {
                ExcelPackage pckg = new ExcelPackage(new FileInfo(file));
                var worksheet = pckg.Workbook.Worksheets[1];
                masterPackage.Workbook.Worksheets.Add("Sheet1");
                var finalSheet = masterPackage.Workbook.Worksheets[1];
                int rowcount = worksheet.Dimension.End.Row;
                int colcount = worksheet.Dimension.End.Column;
                int resultFinalRow = finalSheet.Dimension.End.Row;//apparently getting set to null
                int resultFinalCol = finalSheet.Dimension.End.Column;//apparently getting set to null

                if (resultFinalRow == 0 && resultFinalCol == 0)
                {
                    worksheet.Cells[1, colcount].Copy(finalSheet.Cells[1, resultFinalCol]);
                    worksheet.Cells[1, 2, rowcount, colcount].Copy(finalSheet.Cells[2, resultFinalCol]);
                }
                else
                    worksheet.Cells[1, 2, rowcount, colcount].Copy(finalSheet.Cells[resultFinalRow, resultFinalCol]);
                

                //foreach (var sheet in pckg.Workbook.Worksheets) //-this section puts the worksheets into a single .xlsx file
                //{
                    //check name of worksheet, in case that worksheet with same name already exist exception will be thrown by EPPlus
                    //string workSheetName = sheet.Name;
                    //foreach (var masterSheet in masterPackage.Workbook.Worksheets)
                    //{
                    //    if (sheet.Name == masterSheet.Name)
                    //    {
                    //        workSheetName = string.Format("{0}_{1}", workSheetName, DateTime.Now.ToString("yyyyMMddhhssmmm"));
                    //    }
                    //}

                    //add new sheet
                    //masterPackage.Workbook.Worksheets.Add(workSheetName, sheet);
                //}
            }



            masterPackage.SaveAs(new FileInfo(resultFile));
            Console.WriteLine("File created");
        }

    }
}
