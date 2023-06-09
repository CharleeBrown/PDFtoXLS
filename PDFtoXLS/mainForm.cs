using Aspose.Pdf;
using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using Aspose.Cells;


namespace PDFtoXLS
{
    public partial class mainForm : Form
    {
        PullFiles pullfiles = new PullFiles();
        public mainForm()
        {
            InitializeComponent();

            // Setting listView details and column attributes
            documentListView.View = View.Details;
            documentListView.Columns.Add("File Path");
            documentListView.Columns[0].Width = 175;
            documentListView.Columns.Add("Filename");
            documentListView.Columns[1].Width = 150;
        }
   
        private void button1_Click(object sender, EventArgs e)
        {
            pullfiles.ObtainPDF(documentListView, defaultNamesCheck);
            
        }
        private void button2_Click(object sender, EventArgs e)
        {
            // ReadStream
            Stream newStream;

            // FileDialog to select PDFs
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                // FileDialog settings. Enabling MultiSelect and adding Filter for only PDFs
                dlg.Multiselect = true;
                dlg.Filter = "PDF (*.pdf)|*.pdf|All Files (*.*)|*.*";

                // If the list has items
                if (documentListView.Items.Count > 0)
                {
                    // Clear the items, update the list, and refresh the list
                    documentListView.Items.Clear();
                    documentListView.Update();
                    documentListView.Refresh();
                }

                // If the file selection is OK
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    foreach (String files in dlg.FileNames)
                    {
                        try
                        {
                            // If the file is not null
                            if ((newStream = dlg.OpenFile()) != null)
                            {
                                using (newStream)
                                {
                                    // The stream is used to add the file name to the 
                                    FileInfo info = new FileInfo(files);
                                    ListViewItem gets = new ListViewItem(files);
                                    gets.SubItems.Add(info.Name);
                                    //fileBox.Items.Add(files);
                                    documentListView.Items.Add(gets);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("File Error" + ex.Message);
                        }
                    }

                }
            }

        }
       

        private void Form1_Load(object sender, EventArgs e)
        {

        }


    }

    public class ExcelClear
    {


        public void CleanXls(string filename)
        {
            try
            {
                //Excel instance
                Excel.Application xlApp = new Excel.Application();

                //Opening the recently saved workbook
                Excel.Workbook wbk = xlApp.Workbooks.Open(filename);

                //Setting the variable to the activesheet
                Excel.Worksheet wks = wbk.ActiveSheet;

                // For each image in the worksheet
                foreach (Excel.Shape sh in wks.Shapes)
                {
                    // Delete the image
                    sh.Delete();
                }

                //Setting Range variables for A1 and A2
                Excel.Range miscRange = wks.Range["A1"];
                Excel.Range twoRange = wks.Range["A2"];

                // Unmerging the ranges
                miscRange.UnMerge();
                twoRange.UnMerge();

                // Clearing the ranges
                miscRange.Clear();
                twoRange.Clear();
                // Fitting the columns
                wks.Columns.AutoFit();
                //Fitting the rows
                wks.Rows.AutoFit();

                //Save and Exit
                wbk.Save();
                wbk.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }


        }
    }
}