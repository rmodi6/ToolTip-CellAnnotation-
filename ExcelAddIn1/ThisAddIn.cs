using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {
        private TextBox txtBox = new TextBox();
    
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //this.Application.SheetSelectionChange += new Excel.AppEvents_SheetSelectionChangeEventHandler(SheetSelectionChange);

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        void Application_WorkbookBeforeSave(Microsoft.Office.Interop.Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
            Excel.Range firstRow = activeWorksheet.get_Range("A1");
            firstRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            Excel.Range newFirstRow = activeWorksheet.get_Range("A1");
            newFirstRow.Value2 = "This text was added by using code";
        }
        Microsoft.Office.Tools.Excel.NamedRange selectedEventRange;
        private void SheetSelectionChange(object sh, Microsoft.Office.Interop.Excel.Range target)
        {
            Excel.Worksheet sheet = (Excel.Worksheet)sh;

            String cellResult = "";
            foreach (Excel.Range c in target.Cells)
            {
                string changedCell = c.get_Address(Type.Missing, Type.Missing, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                /*DialogResult result=MessageBox.Show("Address:" + changedCell + " Value: " + System.Drawing.ColorTranslator.FromOle((int)((double)c.Interior.Color)), "Save", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes) {
                    MessageBox.Show("Working");
                }*/
                //System.Drawing.Color col = System.Drawing.ColorTranslator.FromOle((int)((double)r.Interior.Color));
                //cellResult = cellResult + "Address:" + changedCell + " Value: " + System.Drawing.ColorTranslator.FromOle((int)((double)c.Interior.Color))+ "\n";
                cellResult = cellResult + "Address:" + changedCell + " Border Info : " + c.Borders.LineStyle +", "+ c.Borders.Weight + ", Font Name: "+ c.Font.Name + ", Style :"+c.Style.Font.Size +", range formula :"+c.HasFormula+"\n";

            }
            DialogResult result = MessageBox.Show(cellResult, "Save", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                CreateDocument(cellResult);
            }

        }

        string[] ConvertToStringArray(System.Array values)
        {

            // create a new string array
            string[] theArray = new string[values.Length];

            // loop through the 2-D System.Array and populate the 1-D String Array
            for (int i = 1; i <= values.Length; i++)
            {
                if (values.GetValue(1, i) == null)
                    theArray[i - 1] = "";
                else
                    theArray[i - 1] = (string)values.GetValue(1, i).ToString();
            }

            return theArray;
        }
        private void WorkbookSheetSelectionChange()
        {
            
        }

        void ThisWorkbook_SheetSelectionChange(object Sh,
            Excel.Range Target)
        {
            Excel.Worksheet sheet = (Excel.Worksheet)Sh;

            this.Application.StatusBar = sheet.Name + ":" +
                Target.get_Address(
                Excel.XlReferenceStyle.xlA1);
        }

        private void CreateDocument(String cellResult)
        {
            try
            {
                //Create an instance for word app  
                Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();

                //Set animation status for word application  
                winword.ShowAnimation = false;

                //Set status for word application is to be visible or not.  
                winword.Visible = false;

                //Create a missing variable for missing value  
                object missing = System.Reflection.Missing.Value;

                //Create a new document  
                Microsoft.Office.Interop.Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);

                //Add header into the document  
                /*foreach (Microsoft.Office.Interop.Word.Section section in document.Sections)
                {
                    //Get the header range and add the header details.  
                    Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                    headerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    headerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlue;
                    headerRange.Font.Size = 10;
                    headerRange.Text = "Header text goes here";
                }

                //Add the footers into the document  
                foreach (Microsoft.Office.Interop.Word.Section wordSection in document.Sections)
                {
                    //Get the footer range and add the footer details.  
                    Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkRed;
                    footerRange.Font.Size = 10;
                    footerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    footerRange.Text = "Footer text goes here";
                }*/

                //adding text to document  
                document.Content.SetRange(0, 0);
                document.Content.Text = cellResult;

                //Save the document  
                object filename = @"E:\temp1.docx";
                document.SaveAs2(ref filename);
                document.Close(ref missing, ref missing, ref missing);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
                MessageBox.Show("Document created successfully !");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        #endregion
    }
}
