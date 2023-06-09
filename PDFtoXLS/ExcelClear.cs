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
    internal class ExcelClear
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
