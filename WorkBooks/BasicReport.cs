﻿using System;
using System.Collections.Generic;
using OfficeOpenXml;
using ServiceNow_Report_Assistant.Enums;
using System.IO;

namespace ServiceNow_Report_Assistant.WorkBooks
{
    public abstract class BasicReport
    {
        public string Name { get; set; }
        public string Path { get; set; }
        public WorkBookType Type { get; set; }

        public BasicReport(string name, string path)
        {
            Name = name;
            Path = Path;
        }

        public virtual void ReadReport(string path)
        {
            FileInfo existingFile = new FileInfo(path);

            using (ExcelPackage excel = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = excel.Workbook.Worksheets[1];
            }
        }
    }
}
