using Aspose.Cells;
using System;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace PDFtoXLS
{
    internal class Clearing
    {
        public void PMICleanXls(string filename)
        {
            //Creating an Excel Application instance
                Excel.Application xlApp = new Excel.Application();
                xlApp.Visible = false;
                xlApp.DisplayAlerts= false;

            //Creating a Workbook instance
                Excel.Workbook wbk = null;

            try
            {
                //Opening the requested filename
                    wbk = xlApp.Workbooks.Open(filename);

                //Creating a worksheet variable
                    Excel.Worksheet wks = wbk.ActiveSheet;

                // For each image in the worksheet
                    foreach (Excel.Shape sh in wks.Shapes)
                    {
                          // Move the image
                        sh.IncrementLeft(148);
                        sh.IncrementTop(5);
                    }

                    //Unmerge A2 and merge it with G2 and center the text
                    wks.Range["A2"].UnMerge();
                    wks.Range["A2", "G2"].Merge();
                    wks.Range["A2"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    //Unmerge A3 and then merge the range A3 to G3                    
                    wks.Range["A3"].UnMerge();
                    Excel.Range borderRange = wks.Range["A3", "G3"];
                    borderRange.Merge();
               
                    borderRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    borderRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 2d;

                    //Setting Range variables for A1 and A2
                    // These ranges hold certain symbols and counts from the certificate
                    Excel.Range miscRange = wks.Range["A1", "C1"];
                    Excel.Range sigmaRange = wks.Range["C19"];
                    Excel.Range splitSigmaRange = wks.Range["D19"];

                    Excel.Range countRange = wks.Range["B19", "B44"];
                    Excel.Range symbolRange = wks.Range["C19", "C44"];
                    Excel.Range otherRange = wks.Range["D19", "D44"];

                    sigmaRange.UnMerge();
                    sigmaRange.Clear();
                    sigmaRange.Value = "+/-";
                    
                    // Generating the character symbols at the top of the column
                    string mainVal = Char.ConvertFromUtf32(0x00B1) + " 2"+ Char.ConvertFromUtf32(0x0073);

                    // Inserting the characters into the ranges.
                    splitSigmaRange.Value = mainVal;
                    splitSigmaRange.Font.Name = "Symbol";
                   
                    // Setting the column widths
                    wks.Columns["A:A"].ColumnWidth = 14.00;
                    wks.Columns["B:B"].ColumnWidth = 13.89;
                    wks.Columns["C:C"].ColumnWidth = 8.10;
                    wks.Columns["D:D"].ColumnWidth = 8.00;

                    // Aligning the columns to the right.
                    countRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    symbolRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    otherRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

                    // Setting the field information.
                    wks.Range["A13"].Value = "PART #";
                    wks.Range["A14"].UnMerge();
                    wks.Range["A14"].Value = "JOB #";

                    // Merging the cells 
                    wks.Range["A13", "C13"].Merge();
                    wks.Range["A14", "C14"].Merge();
                    wks.Range["A45"].UnMerge();
                    wks.Range["A45", "D45"].Merge();

                    // Setting the row height
                    wks.Range["A45"].RowHeight = 42;

                    // Unmerging the ranges
                    miscRange.UnMerge();

                    // Clearing the ranges
                    miscRange.Clear();
                    wbk.Author = "";
                    //Save and Exit
                    wbk.Save();
                    wbk.Close();
                }

                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    wbk.Close();
                    return;
                }
        }

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
                    Excel.Range miscRange = wks.Range["A1", "C1"];

                    // Unmerging the ranges
                    miscRange.UnMerge();


                    // Clearing the ranges
                    miscRange.Clear();

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

