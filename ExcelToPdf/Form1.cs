using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;


//Ref : https://medium.com/better-programming/convert-excel-files-into-pdf-in-c-net-5566f170a70e

namespace ExcelToPdf
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            
            InitializeComponent();
            lblMsg.Hide();
        }

        private void btnDestination_Click(object sender, EventArgs e)
        {

            FolderBrowserDialog fbdDestination = new FolderBrowserDialog() {
                ShowNewFolderButton=true
            };

            fbdDestination.Description = "Select Destination Folder";
            fbdDestination.ShowDialog();
            tbDestination.Text = fbdDestination.SelectedPath;
            
        }

        private void btnSource_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbdSource = new FolderBrowserDialog()
            {
                ShowNewFolderButton = true
            };

            fbdSource.Description = "Select Source Folder";
            fbdSource.ShowDialog();
            tbSource.Text=fbdSource.SelectedPath;
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            string directoryWithExcelFiles;
            string sourceFolder = tbSource.Text;
            string destinationFolder = tbDestination.Text;
            string sMsg = @"Sit back and Relax, Your Excel(s) are being Converted to PDF...";

            tbDestination.Hide();
            tbSource.Hide();
            btnConvert.Hide();
            btnSource.Hide();
            btnDestination.Hide();
            lblMsg.Text = sMsg;
            lblMsg.Show();

            if (sourceFolder.Length == 0)
            {
                // If no directory path is passed as argument, consider the current process directory
                //directoryWithExcelFiles = Directory.GetCurrentDirectory();
                MessageBox.Show("Select Source Folder.", "Select Folder");
            }
            else
            {
                directoryWithExcelFiles = Path.GetFullPath(sourceFolder);

                if (destinationFolder.Length ==0)
                {
                    destinationFolder = sourceFolder;
                }

                var excelFilesToConvert = Directory.EnumerateFiles(directoryWithExcelFiles, "*.xls");
                var excelInteropExcelToPdfConverter = new ExcelInteropExcelToPdfConverter();

                try
                {
                    excelInteropExcelToPdfConverter.ConvertToPdf(excelFilesToConvert, destinationFolder);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Something went wrong: {ex.Message}");
                    Environment.ExitCode = -1;

                    tbDestination.Show();
                    tbSource.Show();
                    btnConvert.Show();
                    btnSource.Show();
                    btnDestination.Show();
                    lblMsg.Text = sMsg;
                    lblMsg.Hide();

                    return;
                }
            }

            MessageBox.Show("Operation completed");
        }

        private void GetAllFiles()
        {

        }
    }


    //HELPER CLASS
    public class ExcelApplicationWrapper : IDisposable
    {
        public Microsoft.Office.Interop.Excel.Application ExcelApplication { get; }

        public ExcelApplicationWrapper()
        {
            ExcelApplication = new Microsoft.Office.Interop.Excel.Application();
        }

        public void Dispose()
        {
            // Each file I open is locked by the background EXCEL.exe until it is quitted
            ExcelApplication.Quit();
            Marshal.ReleaseComObject(ExcelApplication);
        }
    }

    //THE CONVERTER
    public class ExcelInteropExcelToPdfConverter
    {
        public void ConvertToPdf(IEnumerable<string> excelFilesPathToConvert,string dest)
        {
            using (var excelApplication = new ExcelApplicationWrapper())
            {
                foreach (var excelFilePath in excelFilesPathToConvert)
                {
                    var thisFileWorkbook = excelApplication.ExcelApplication.Workbooks.Open(excelFilePath);
                    string newPdfFilePath = Path.Combine(
                        (dest.Length == 0 ? Path.GetDirectoryName(excelFilePath):dest),
                        $"{Path.GetFileNameWithoutExtension(excelFilePath)}.pdf");

                    thisFileWorkbook.ExportAsFixedFormat(
                        Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
                        newPdfFilePath);
                 
                    thisFileWorkbook.Close(false, excelFilePath);
                    Marshal.ReleaseComObject(thisFileWorkbook);
                }
            }
        }
    }


}
