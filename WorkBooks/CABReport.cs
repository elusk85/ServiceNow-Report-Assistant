using System;
using System.Collections.Generic;
using OfficeOpenXml;
using ServiceNow_Report_Assistant.Enums;
using System.IO;

namespace ServiceNow_Report_Assistant.WorkBooks
{
    public class CABReport:BasicReport
    {
            public CABReport(string name, string path) : base(name, path)
            {
                Name = name;
                Path = Path;

                Console.WriteLine("Name: " + Name + "Path: " + Path);
            }

            public override void ReadReport(string path)
            {
                FileInfo existingFile = new FileInfo(path);

                using (ExcelPackage excel = new ExcelPackage(existingFile))
                {
                    ExcelWorksheet worksheet = excel.Workbook.Worksheets[1];
                }
            }

     }
 }
