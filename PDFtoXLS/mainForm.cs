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
            FileRead read = new FileRead();
            read.OpenFile(documentListView);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}