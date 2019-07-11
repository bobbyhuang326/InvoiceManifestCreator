using System;
//using WordEdit;
using System.IO;
using Manifest;

namespace ManipulateOffice
{
    class Program
    {
        static void Main(string[] args)
        {
            ReportGenerate r = new ReportGenerate();
            r.manifestFileName = "广美20190709.xls";
            r.manifestFileDir = @"C:\Users\Bobby\Desktop\清单模板";
            r.templateFileName = "template.xlsx";
            r.templateFileDir = @"C:\Users\Bobby\Desktop\清单模板";
            r.ExcelGenearte();
        }
    }
}
