using Microsoft.WindowsAPICodePack.Dialogs;
using SSMSTextResultExcelLib;
using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace SSMSTextResultExcel
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }




        private void btnClear_Click(object sender, EventArgs e)
        {
            //txtSQLInput.Text = txtSQLOutput.Text = lblWorkResult.Text = "";
        }

        

    

        private void btnOpenFolder_Click(object sender, EventArgs e)
        {
            var dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                txtFolderPath.Text = dialog.FileName;
            }
        }

        private void btnExcelMakeAuto_Click(object sender, EventArgs e)
        {
            String FolderName = txtFolderPath.Text;
            string outputFolderName = "excelOutput";

            if (string.IsNullOrEmpty(FolderName))
            {
                MessageBox.Show("폴더 경로 입력");
                return;
            }

            // Output 폴더 만들기ddd
            DirectoryInfo outPutFolder = new DirectoryInfo(FolderName + $"\\{outputFolderName}");
            if (outPutFolder.Exists == false)
                outPutFolder.Create();

            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(FolderName);         
        
            foreach (System.IO.FileInfo file in di.GetFiles())
            {
                if (file.Extension.ToLower().CompareTo(".txt") == 0)
                {
                    String FileNameOnly = file.Name.Substring(0, file.Name.Length - 4);
                    String fullFileName = file.FullName;

                    btnClear_Click(sender, e);
                    string str = File.ReadAllText(fullFileName, Encoding.Default);

                    //////////////////////////////////////////////////////////
                    var excelResult = new ExportFile();

                    try
                    {
                        excelResult.ExcelResult(str, file.DirectoryName + $"\\{outputFolderName}\\" + FileNameOnly + ".xlsx");
                        //txtSQLOutput.Text = "Excel Output finished";
                    } catch (IndexOutOfRangeException err)
                    {
                        //txtSQLOutput.Text = err.Message;
                    } catch (ArgumentOutOfRangeException)
                    {
                        //txtSQLOutput.Text = "Input String format is invalid";
                    } catch (InvalidOperationException)
                    {
                        //txtSQLOutput.Text = "Excel File is opened. close";
                    }
                }
            }


            
        }
    }
}
