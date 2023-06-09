using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System;
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

    public class PullFiles
    {
        public void ObtainPDF(ListView docList, CheckBox defaultNames)
        {
            // Instance of FolderBrowser
            FolderBrowserDialog fldr = new FolderBrowserDialog();

            // Folder Browser description
            fldr.Description = "Choose Folder to Save";

            // If the folder selected is viable
            if (fldr.ShowDialog() == DialogResult.OK)
            {
                // For each PDF listed
                foreach (ListViewItem link in docList.Items)
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
                    if (defaultNames.Checked == false)
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
    } }
