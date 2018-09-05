using System;
using System.Collections.Generic;
using OfficeOpenXml;
using ServiceNow_Report_Assistant.Enums;

namespace ServiceNow_Report_Assistant.WorkBooks
{
    public abstract class AssetReport
    {
        public string Name { get; set; }
        public string Path { get; set; }
        public ExcelPackage excel = new ExcelPackage();
        public WorkBookType Type { get; set; }

        public AssetReport(string name, string path)
        {
            Name = name;
            Path = Path;
        }

    }
}
