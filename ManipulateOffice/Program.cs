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
            var form = new ExcelUploadForm(r);
            form.OpenFileDialog();
            Application.EnableVisualStyles();
            Application.Run(form);
        }
    }
}
