using System;
//using WordEdit;
using System.IO;
using System.Windows.Forms;
using Manifest;

namespace ManipulateOffice
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            ReportGenerate r = new ReportGenerate();
            r.ManifestFileName = "20191124.xlsx";
            r.ManifestFileDir = "C:\\Users\\Bobby\\Desktop";
            r.ExcelGenearte();
//            var form = new ExcelUploadForm(r);
//            form.OpenFileDialog();+
//            Application.EnableVisualStyles();
//            Application.Run(form);
        }
    }
}
