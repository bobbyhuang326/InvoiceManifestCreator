using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ManipulateOffice
{
    //class SimpleUI
    //{
    //    private string fileName { get; set; }
    //    private  string filePath { get; set; }
    //    //private void buttonSetOutputPath(object sender, EventArgs e) {
            
    //    //}
    //}

    class ExcelUploadForm : Form {
        private OpenFileDialog op;
        private Button[] browseBtn;
        public string fileName { get; set; }
        public string filePath { get; set; }

        public ExcelUploadForm() {
            //op = new OpenFileDialog();
            browseBtn = new Button[4];
        }

        public void OpenFileDialog() {
            //op.ShowDialog(); 
            //browseBtn.
            //browseBtn.Click += op.;// attatch openFileDialog to browseBtn
            //op.FileOk += buttonGetImportFile; // attach file path
            browseBtn[0] = new Button();
            browseBtn[1] = new Button();
            browseBtn[2] = new Button();
            browseBtn[3] = new Button();
            browseBtn[0].Text = "清单文件上传";
            browseBtn[1].Text = "清单模板上传";
            browseBtn[1].Location = new System.Drawing.Point(browseBtn[0].Left,browseBtn[0].Bottom+10);
            browseBtn[2].Text = "输出地址选择";
            browseBtn[3].Text = "导出";
            this.Controls.AddRange(browseBtn);
        }

        //private void attachEvent() {
        //    this.op. += new EventHandler(buttonGetImportFile);
        //}

        private void buttonGetImportFile(object sender, EventArgs e) // attach this to open button in openFileDialog
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel files (*.xls,*.xlsx)"; // file types, that will be allowed to upload
            dialog.Multiselect = false; // allow/deny user to upload more than one file at a time
            if (dialog.ShowDialog() == DialogResult.OK) // if user clicked OK
            {
                this.fileName = Path.GetFileName(dialog.FileName); // get name of file
                this.filePath = Path.GetDirectoryName(dialog.FileName); // get path of file
            }
        }
    } 




}
