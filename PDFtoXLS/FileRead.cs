using Aspose.Pdf;
using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;


namespace PDFtoXLS
{
    internal class FileRead
    {

        public void OpenFile(ListView documentListView)
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
    }
}
