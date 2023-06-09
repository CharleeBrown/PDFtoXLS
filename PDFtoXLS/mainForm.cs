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
        public mainForm()
        {
            InitializeComponent();

            // Setting listView details and column attributes
            listView1.View = View.Details;
            listView1.Columns.Add("File Path");
            listView1.Columns[0].Width = 175;
            listView1.Columns.Add("Filename");
            listView1.Columns[1].Width = 150;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Instance of FolderBrowser
            FolderBrowserDialog fldr = new FolderBrowserDialog();

            // Folder Browser description
            fldr.Description = "Choose Folder to Save";

            // If the folder selected is viable
            if (fldr.ShowDialog() == DialogResult.OK)
            {
                // For each PDF listed
                foreach (ListViewItem link in listView1.Items)
                {
                    // The file info is obtained
                    FileInfo nfo = new FileInfo(link.Text);

                    //The PDF read VIA the path
                    Document pdfs = new Document(link.Text);

                    // Save options for the Aspose Package
                    ExcelSaveOptions options = new ExcelSaveOptions();

                    //variable for filename length
                    int nameLength = nfo.Name.Length;

                    //The textbox for the new filename will give the filename minus the extension type
                    String strs = nfo.Name.Substring(0, nameLength - 4);

                    // If the Default Filenames checkbox is unchecked. Ask for each filename.
                    if (defaultNamesCheck.Checked == false)
                    {
                        //If the filename is valid
                        if (InputBox("PDFtoXLS", "Enter new Filename", ref strs) == DialogResult.OK)
                        {
                            //The save folder path is saved
                            string paths = fldr.SelectedPath;

                            //The path is combined and the proper extension added
                            string newPlace = Path.Combine(paths, strs + ".xlsx");

                            //The PDF is converted and saved under the new extension.
                            pdfs.Save(newPlace, options);

                            // An instance of the ExcelClear class is created.
                            ExcelClear clear = new ExcelClear();

                            // The "CleanXls" method is run on the newly saved file.
                            clear.CleanXls(newPlace);
                        }
                    }
                    else
                    {
                        //The save folder path is saved
                        string paths = fldr.SelectedPath;

                        //The path is combined and the proper extension added
                        string newPlace = Path.Combine(paths, strs + ".xlsx");

                        //The PDF is converted and saved under the new extension.
                        pdfs.Save(newPlace, options);

                        // An instance of the ExcelClear class is created.
                        ExcelClear clear = new ExcelClear();

                        // The "CleanXls" method is run on the newly saved file.
                        clear.CleanXls(newPlace);
                    }
                }
            }
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
                if (listView1.Items.Count > 0)
                {
                    // Clear the items, update the list, and refresh the list
                    listView1.Items.Clear();
                    listView1.Update();
                    listView1.Refresh();
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
                                    listView1.Items.Add(gets);
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
        public static DialogResult InputBox(string title, string promptText, ref string value)
        {

            // Example from https://www.csharp-examples.net/inputbox/ 
            Form form = new Form();
            System.Windows.Forms.Label label = new System.Windows.Forms.Label();
            System.Windows.Forms.TextBox textBox = new System.Windows.Forms.TextBox();
            System.Windows.Forms.Button buttonOk = new System.Windows.Forms.Button();
            System.Windows.Forms.Button buttonCancel = new System.Windows.Forms.Button();

            form.Text = title;
            label.Text = promptText;
            textBox.Text = value;

            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 20, 372, 13);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(396, 107);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            DialogResult dialogResult = form.ShowDialog();
            value = textBox.Text;
            return dialogResult;
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