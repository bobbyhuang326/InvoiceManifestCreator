using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Manifest;

namespace ManipulateOffice
{
    class ExcelUploadForm : Form
    {
        private OpenFileDialog openFileDialog;
        private Button ConfirmButton;
        private Button SelectButton;
        public ReportGenerate _generator { get; set; }

        public ExcelUploadForm(ReportGenerate generator)
        {
            _generator = generator;
        }

        public void OpenFileDialog()
        {
            ConfirmButton = new Button
            {
                Size = new Size(100, 40),
                Location = new Point(15, 120),
                Text = "生成清单"
            };
            SelectButton = new Button
            {
                Size = new Size(100, 40),
                Location = new Point(15, 15),
                Text = "选择文件"
            };
            openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.Multiselect = false;

            SelectButton.Click += SelectButtonClick;
            ConfirmButton.Click += ConfirmButtonClick;
            
            Controls.Add(SelectButton);
            Controls.Add(ConfirmButton);
        }

        private void SelectButtonClick(object sender, EventArgs e) // attach this to open button in openFileDialog
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK) // if user clicked OK
            {
                _generator.ManifestFileName = Path.GetFileName(openFileDialog.FileName); // get name of file
                _generator.ManifestFileDir = Path.GetDirectoryName(openFileDialog.FileName); // get path of file
            }
        }

        private void ConfirmButtonClick(object sender, EventArgs e)
        {
            _generator.ExcelGenearte();
        }
    }
}