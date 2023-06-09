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
        //Generating class instances
        PullFiles pullfiles = new PullFiles();
        FileRead read = new FileRead();
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
            // Function from the PullFiles class that takes a ListView and CheckBox as inputs
            pullfiles.ObtainPDF(documentListView, defaultNamesCheck);
            
        }
        private void button2_Click(object sender, EventArgs e)
        {
           // Function from FileRead that takes a ListView as input
            read.OpenFile(documentListView);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}