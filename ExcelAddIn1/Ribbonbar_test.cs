using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using System.Data;
using Word = Microsoft.Office.Interop.Word;
using System.Drawing;

namespace ExcelAddIn1
{

    public partial class Ribbonbar_test
    {
        //DataGridView grid1 = new DataGridView();
        private void Ribbonbar_test_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void BtnShow_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet sh = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range target = Globals.ThisAddIn.Application.Selection as Excel.Range;
            Excel.Worksheet sheet = (Excel.Worksheet)sh;

            String cellResult = "";
            cellResult = cellResult + "Address: \t"+ "Border Info: \t"  + "Font Name: \t" + "Style: \t" + "range formula: \t" + "\n";

            DataTable table = new DataTable();
            DataTable table1 = new DataTable();
            DataTable table2 = new DataTable();

            DataColumn column;
            DataColumn column1;
            DataColumn column2;
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "A";
            column.ReadOnly = false;
            column.Unique = false;
            table.Columns.Add(column);

            column1 = new DataColumn();
            column1.DataType = System.Type.GetType("System.String");
            column1.ColumnName = "A";
            column1.ReadOnly = false;
            column1.Unique = false;
            table1.Columns.Add(column1);

            column2 = new DataColumn();
            column2.DataType = System.Type.GetType("System.String");
            column2.ColumnName = "A";
            column2.ReadOnly = false;
            column2.Unique = false;
            table2.Columns.Add(column2);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "B";
            column.ReadOnly = false;
            column.Unique = false;
            table.Columns.Add(column);

            column1 = new DataColumn();
            column1.DataType = System.Type.GetType("System.String");
            column1.ColumnName = "B";
            column1.ReadOnly = false;
            column1.Unique = false;
            table1.Columns.Add(column1);

            column2 = new DataColumn();
            column2.DataType = System.Type.GetType("System.String");
            column2.ColumnName = "B";
            column2.ReadOnly = false;
            column2.Unique = false;
            table2.Columns.Add(column2);


            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "C";
            column.ReadOnly = false;
            column.Unique = false;
            table.Columns.Add(column);


            column1 = new DataColumn();
            column1.DataType = System.Type.GetType("System.String");
            column1.ColumnName = "C";
            column1.ReadOnly = false;
            column1.Unique = false;
            table1.Columns.Add(column1);

            column2 = new DataColumn();
            column2.DataType = System.Type.GetType("System.String");
            column2.ColumnName = "C";
            column2.ReadOnly = false;
            column2.Unique = false;
            table2.Columns.Add(column2);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "D";
            column.ReadOnly = false;
            column.Unique = false;
            table.Columns.Add(column);

            column1 = new DataColumn();
            column1.DataType = System.Type.GetType("System.String");
            column1.ColumnName = "D";
            column1.ReadOnly = false;
            column1.Unique = false;
            table1.Columns.Add(column1);

            column2 = new DataColumn();
            column2.DataType = System.Type.GetType("System.String");
            column2.ColumnName = "D";
            column2.ReadOnly = false;
            column2.Unique = false;
            table2.Columns.Add(column2);


            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "E";
            column.ReadOnly = false;
            column.Unique = false;
            table.Columns.Add(column);

            column1 = new DataColumn();
            column1.DataType = System.Type.GetType("System.String");
            column1.ColumnName = "E";
            column1.ReadOnly = false;
            column1.Unique = false;
            table1.Columns.Add(column1);

            column2 = new DataColumn();
            column2.DataType = System.Type.GetType("System.String");
            column2.ColumnName = "E";
            column2.ReadOnly = false;
            column2.Unique = false;
            table2.Columns.Add(column2);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "F";
            column.ReadOnly = false;
            column.Unique = false;
            table.Columns.Add(column);


            column1 = new DataColumn();
            column1.DataType = System.Type.GetType("System.String");
            column1.ColumnName = "F";
            column1.ReadOnly = false;
            column1.Unique = false;
            table1.Columns.Add(column1);

            column2 = new DataColumn();
            column2.DataType = System.Type.GetType("System.String");
            column2.ColumnName = "F";
            column2.ReadOnly = false;
            column2.Unique = false;
            table2.Columns.Add(column2);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "G";
            column.ReadOnly = false;
            column.Unique = false;
            table.Columns.Add(column);


            column1 = new DataColumn();
            column1.DataType = System.Type.GetType("System.String");
            column1.ColumnName = "G";
            column1.ReadOnly = false;
            column1.Unique = false;
            table1.Columns.Add(column1);

            column2 = new DataColumn();
            column2.DataType = System.Type.GetType("System.String");
            column2.ColumnName = "G";
            column2.ReadOnly = false;
            column2.Unique = false;
            table2.Columns.Add(column2);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "H";
            column.ReadOnly = false;
            column.Unique = false;
            table.Columns.Add(column);


            column1 = new DataColumn();
            column1.DataType = System.Type.GetType("System.String");
            column1.ColumnName = "H";
            column1.ReadOnly = false;
            column1.Unique = false;
            table1.Columns.Add(column1);

            column2 = new DataColumn();
            column2.DataType = System.Type.GetType("System.String");
            column2.ColumnName = "H";
            column2.ReadOnly = false;
            column2.Unique = false;
            table2.Columns.Add(column2);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "I";
            column.ReadOnly = false;
            column.Unique = false;
            table.Columns.Add(column);


            column1 = new DataColumn();
            column1.DataType = System.Type.GetType("System.String");
            column1.ColumnName = "I";
            column1.ReadOnly = false;
            column1.Unique = false;
            table1.Columns.Add(column1);

            column2 = new DataColumn();
            column2.DataType = System.Type.GetType("System.String");
            column2.ColumnName = "I";
            column2.ReadOnly = false;
            column2.Unique = false;
            table2.Columns.Add(column2);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "J";
            column.ReadOnly = false;
            column.Unique = false;
            table.Columns.Add(column);

            column1 = new DataColumn();
            column1.DataType = System.Type.GetType("System.String");
            column1.ColumnName = "J";
            column1.ReadOnly = false;
            column1.Unique = false;
            table1.Columns.Add(column1);

            column2 = new DataColumn();
            column2.DataType = System.Type.GetType("System.String");
            column2.ColumnName = "J";
            column2.ReadOnly = false;
            column2.Unique = false;
            table2.Columns.Add(column2);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "K";
            column.ReadOnly = false;
            column.Unique = false;
            table.Columns.Add(column);


            column1 = new DataColumn();
            column1.DataType = System.Type.GetType("System.String");
            column1.ColumnName = "K";
            column1.ReadOnly = false;
            column1.Unique = false;
            table1.Columns.Add(column1);

            column2 = new DataColumn();
            column2.DataType = System.Type.GetType("System.String");
            column2.ColumnName = "K";
            column2.ReadOnly = false;
            column2.Unique = false;
            table2.Columns.Add(column2);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "L";
            column.ReadOnly = false;
            column.Unique = false;
            table.Columns.Add(column);


            column1 = new DataColumn();
            column1.DataType = System.Type.GetType("System.String");
            column1.ColumnName = "L";
            column1.ReadOnly = false;
            column1.Unique = false;
            table1.Columns.Add(column1);

            column2 = new DataColumn();
            column2.DataType = System.Type.GetType("System.String");
            column2.ColumnName = "L";
            column2.ReadOnly = false;
            column2.Unique = false;
            table2.Columns.Add(column2);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "M";
            column.ReadOnly = false;
            column.Unique = false;
            table.Columns.Add(column);


            column1 = new DataColumn();
            column1.DataType = System.Type.GetType("System.String");
            column1.ColumnName = "M";
            column1.ReadOnly = false;
            column1.Unique = false;
            table1.Columns.Add(column1);

            column2 = new DataColumn();
            column2.DataType = System.Type.GetType("System.String");
            column2.ColumnName = "M";
            column2.ReadOnly = false;
            column2.Unique = false;
            table2.Columns.Add(column2);


            DataRow workRow;
            DataRow workRow1;
            DataRow workRow2;
            //// new code ended

            string nl = Environment.NewLine;
            string del = "@#@";
            int maxRow = 0;

            foreach (Excel.Range c in target.Cells)
            {
                string changedCell = c.get_Address(Type.Missing, Type.Missing, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                string[] str = c.Address.Split('$');
                int rowNo = Int32.Parse(str[2]);                

                while(maxRow < rowNo)
                {
                    workRow = table.NewRow();
                    workRow1 = table1.NewRow();
                    workRow2 = table2.NewRow();
                    table.Rows.Add(workRow);
                    table1.Rows.Add(workRow1);
                    table2.Rows.Add(workRow2);
                    maxRow++;
                }

                            
                string a = "";
                try
                {
                    if (c.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle != -4142)
                    {
                        a = a + " " + "Right: "+(Microsoft.Office.Interop.Excel.XlLineStyle)c.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle+"-"+ (Microsoft.Office.Interop.Excel.XlBorderWeight)c.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight + ";";
                    }
                }
                catch (Exception)
                {

                }

                try
                {
                   if(c.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle!=-4142)
                        a = a + " " + "Left: " + (Microsoft.Office.Interop.Excel.XlLineStyle)c.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle + "-" + (Microsoft.Office.Interop.Excel.XlBorderWeight)c.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight + ";";
                }
                catch (Exception)
                {

                }

                try
                {
                    if(c.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle!=-4142)
                        a = a + " " + "Bottom: " + (Microsoft.Office.Interop.Excel.XlLineStyle)c.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle+"-" + (Microsoft.Office.Interop.Excel.XlBorderWeight)c.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight + ";";
                }
                catch (Exception)
                {

                }

                try
                {
                    if(c.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle!=-4142)
                        a = a + " " + "Top: " + (Microsoft.Office.Interop.Excel.XlLineStyle)c.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle+"-" + (Microsoft.Office.Interop.Excel.XlBorderWeight)c.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight;
                }
                catch (Exception)
                {

                }
                String info = "Address: "+changedCell + del + nl +"Width: "+c.EntireColumn.Width +  del + nl + "Height: " + c.EntireRow.Height + del + nl + "Value: " +c.Value + del + nl + "Font Name: "+c.Font.Name + del + nl + "Font Size: "+c.Font.Size + del + nl +
                    "BgColor: "+System.Drawing.ColorTranslator.FromOle((int)((double)c.Interior.Color)) + del + nl + "Font Color: "+(Microsoft.Office.Interop.Excel.XlRgbColor)c.Font.Color + del + nl +
                    "Italics: "+c.Font.Italic + del + nl + "Bold: "+c.Font.Bold + del + nl + "Underline: "+(Microsoft.Office.Interop.Excel.XlUnderlineStyle)c.Font.Underline + del + nl +
                    "Horizontal Alignment: "+(Microsoft.Office.Interop.Excel.XlHAlign)c.HorizontalAlignment + del + nl + "Vertical Alignment: "+(Microsoft.Office.Interop.Excel.XlVAlign)c.VerticalAlignment + del + nl +
                    "Format: "+c.NumberFormat + del + nl + "Formula: " + c.Formula + del + nl + "Border: " + a + del + nl + " ";

                String info1 = "Address: " + changedCell + nl + "Width: " + c.EntireColumn.Width + nl + "Height: " + c.EntireRow.Height + nl + "Value: " + c.Value + nl + "Font Name: " + c.Font.Name + nl + "Font Size: " + c.Font.Size + nl +
                    "BgColor: " + System.Drawing.ColorTranslator.FromOle((int)((double)c.Interior.Color)) + nl + "Font Color: " + (Microsoft.Office.Interop.Excel.XlRgbColor)c.Font.Color + nl +
                    "Italics: " + c.Font.Italic + nl + "Bold: " + c.Font.Bold + nl + "Underline: " + (Microsoft.Office.Interop.Excel.XlUnderlineStyle)c.Font.Underline + nl +
                    "Horizontal Alignment: " + (Microsoft.Office.Interop.Excel.XlHAlign)c.HorizontalAlignment + nl + "Vertical Alignment: " + (Microsoft.Office.Interop.Excel.XlVAlign)c.VerticalAlignment + nl +
                    "Format: " + c.NumberFormat + nl + "Border: " + a + nl + " ";

                String info2 = "[" + changedCell +"," + c.Value +"," + "Width: " + c.EntireColumn.Width + "," + "Height: " + c.EntireRow.Height + "," + c.Font.Name + "," + c.Font.Size +
                    "," + System.Drawing.ColorTranslator.FromOle((int)((double)c.Interior.Color)) + "," + (Microsoft.Office.Interop.Excel.XlRgbColor)c.Font.Color +
                    "," + c.Font.Italic + "," + c.Font.Bold +"," + (Microsoft.Office.Interop.Excel.XlUnderlineStyle)c.Font.Underline + 
                    "," + (Microsoft.Office.Interop.Excel.XlHAlign)c.HorizontalAlignment + "," + (Microsoft.Office.Interop.Excel.XlVAlign)c.VerticalAlignment +
                    "," + c.NumberFormat +"," + a + "]";

                table.Rows[rowNo-1][str[1]] = info;
                //table1.Rows[rowNo - 1][str[1]] = info1;
                //table2.Rows[rowNo - 1][str[1]] = info2;

            }

            DataGridView grid = new DataGridView();
            DataGridView grid1 = new DataGridView();
            DataGridView grid2 = new DataGridView();

            grid.Size = new System.Drawing.Size(900,600);
            grid.DataSource = table;


            grid1.Size = new System.Drawing.Size(900, 600);
            grid1.DataSource = table1;

            grid2.Size = new System.Drawing.Size(900, 600);
            grid2.DataSource = table2;

            //grid1 = grid;

            grid.DataBindingComplete += new DataGridViewBindingCompleteEventHandler(dgv_DataBindingComplete);
            

            using (Form form = new Form())
            {
                form.Text = "Cell Info";
                form.Size = new System.Drawing.Size(900,600);
                form.Controls.Add(grid);
                //form.Controls.Add(grid1);
                //form.Controls.Add(grid2);

                //form.Controls.Add()
                form.ShowDialog();

                DialogResult result = MessageBox.Show("Do you want the save the results??", "Save", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    //grid1.Size = new System.Drawing.Size(5000, 5000);
                    //Export_Data_To_Word(grid1);
                    //exportToExcel_Click(grid);
                    exportToExcel_Click(grid1);
                    //exportToExcel_Click(grid2);

                }

            }  
        }

        

        void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            string nl = Environment.NewLine;
            string del = "@#@";
            if (e.Value != null && e.RowIndex > -1)
            {
                string content = e.Value.ToString();
                string[] line = content.Split(new string[] { del }, StringSplitOptions.None);
                StringFormat sf = new StringFormat();
                sf.Alignment = StringAlignment.Center;
                sf.LineAlignment = StringAlignment.Center;
                

                e.Paint(e.CellBounds, DataGridViewPaintParts.All & ~DataGridViewPaintParts.ContentForeground);

                SizeF[] size = new SizeF[line.Length];
                for (int i = 0; i < line.Length; ++i)
                {
                    size[i] = e.Graphics.MeasureString(line[i], e.CellStyle.Font);
                }

                RectangleF rec = new RectangleF(e.CellBounds.Location, new Size(0, 0));
                SizeF lastSize = new SizeF(0,0);
                using (SolidBrush bblack = new SolidBrush(Color.Black), bred = new SolidBrush(Color.Red), borange = new SolidBrush(Color.DarkOrange),
                    bgreen = new SolidBrush(Color.Green), bblue = new SolidBrush(Color.Blue), bcyan = new SolidBrush(Color.Cyan),
                    bmagenta = new SolidBrush(Color.Magenta), bgold = new SolidBrush(Color.Gold), bnavy = new SolidBrush(Color.Navy),
                    bfuchsia = new SolidBrush(Color.Fuchsia), bcoral = new SolidBrush(Color.Coral), bindigo = new SolidBrush(Color.Indigo), blime = new SolidBrush(Color.Lime),
                    bfirebrick = new SolidBrush(Color.Firebrick), btomato = new SolidBrush(Color.Tomato))
                {
                    for (int i = 0; i < line.Length; i++)
                    {
                        if (i > 1)
                        {
                            lastSize = new SizeF(size[i-1].Width, e.CellBounds.Height );
                            rec = new RectangleF(new PointF(rec.Location.X + rec.Width - lastSize.Width, rec.Location.Y + 14), new SizeF(size[i].Width, e.CellBounds.Height));
                        } else if (i==0)
                        {
                            rec = new RectangleF(new PointF(rec.Location.X + rec.Width - lastSize.Width, rec.Location.Y - 100), new SizeF(size[i].Width, e.CellBounds.Height));
                        } else
                        {
                            lastSize = new SizeF(size[i - 1].Width, e.CellBounds.Height);
                            rec = new RectangleF(new PointF(rec.Location.X + rec.Width - lastSize.Width, rec.Location.Y + 9), new SizeF(size[i].Width, e.CellBounds.Height));
                        }

                        switch (i) {
                            case 0: e.Graphics.DrawString(line[i], e.CellStyle.Font, bblack, rec, sf);
                                break;
                            case 1:
                                e.Graphics.DrawString(line[i] , e.CellStyle.Font, bred, rec, sf);
                                break;
                            case 2:
                                e.Graphics.DrawString(line[i], e.CellStyle.Font, borange, rec, sf);
                                break;
                            case 3:
                                e.Graphics.DrawString(line[i], e.CellStyle.Font, bgreen, rec, sf);
                                break;
                            case 4:
                                e.Graphics.DrawString(line[i], e.CellStyle.Font, bblue, rec, sf);
                                break;
                            case 5:
                                e.Graphics.DrawString(line[i], e.CellStyle.Font, bcyan, rec, sf);
                                break;
                            case 6:
                                e.Graphics.DrawString(line[i], e.CellStyle.Font, bmagenta, rec, sf);
                                break;
                            case 7:
                                e.Graphics.DrawString(line[i], e.CellStyle.Font, bgold, rec, sf);
                                break;
                            case 8:
                                e.Graphics.DrawString(line[i], e.CellStyle.Font, bnavy, rec, sf);
                                break;
                            case 9:
                                e.Graphics.DrawString(line[i], e.CellStyle.Font, bfuchsia, rec, sf);
                                break;
                            case 10:
                                e.Graphics.DrawString(line[i], e.CellStyle.Font, bcoral, rec, sf);
                                break;
                            case 11:
                                e.Graphics.DrawString(line[i], e.CellStyle.Font, bindigo, rec, sf);
                                break;
                            case 12:
                                e.Graphics.DrawString(line[i], e.CellStyle.Font, blime, rec, sf);
                                break;
                            case 13:
                                e.Graphics.DrawString(line[i], e.CellStyle.Font, bfirebrick, rec, sf);
                                break;
                            default:
                                e.Graphics.DrawString(line[i], e.CellStyle.Font, btomato, rec, sf);
                                break;
                        }
                    }

                }
                e.Handled = true;
            }

        }

        private void dgv_DataBindingComplete(Object sender, DataGridViewBindingCompleteEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;
            
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

            foreach (DataGridViewColumn col in dgv.Columns)
            {
                col.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            }

            foreach (DataGridViewRow row in dgv.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {

                    dgv.CellPainting += new DataGridViewCellPaintingEventHandler(dataGridView1_CellPainting);
                }
            }


        }

        private void exportToExcel_Click(DataGridView transcationTableDataGridView)
        {
            try
            {
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                app.Visible = true;
                worksheet = workbook.Sheets["Sheet1"];
                worksheet = workbook.ActiveSheet;
                worksheet.Name = "Records";

                try
                {
                    /*for (int i = 0; i < transcationTableDataGridView.Columns.Count; i++)
                    {
                        worksheet.Cells[1, i + 1] = transcationTableDataGridView.Columns[i].HeaderText;
                    }*/
                    for (int i = 0; i < transcationTableDataGridView.Rows.Count; i++)
                    {
                        for (int j = 0; j < transcationTableDataGridView.Columns.Count; j++)
                        {
                            if (transcationTableDataGridView.Rows[i].Cells[j].Value != null && !transcationTableDataGridView.Rows[i].Cells[j].Value.ToString().Equals(""))
                            {
                                worksheet.Cells[i + 1, j + 1] = transcationTableDataGridView.Rows[i].Cells[j].Value.ToString();
                                
                                /*Excel.Range ColorMeMine= worksheet.Cells[i + 1, j + 1];
                                int ind = -1,ptr=0,color;
                                foreach (var index in transcationTableDataGridView.Rows[i].Cells[j].Value.ToString().findAll(Environment.NewLine))
                                {
                                    switch (ptr) {
                                        case 0: color=System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                                            break;
                                        case 1:
                                            color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                            break;
                                        case 2:
                                            color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                                            break;
                                        case 3:
                                            color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                                            break;
                                        case 4:
                                            color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                                            break;
                                        case 5:
                                            color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Cyan);
                                            break;
                                        case 6:
                                            color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Magenta);
                                            break;
                                        case 7:
                                            color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gold);
                                            break;
                                        case 8:
                                            color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Navy);
                                            break;
                                        case 9:
                                            color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Fuchsia);
                                            break;
                                        case 10:
                                            color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Coral);
                                            break;
                                        case 11:
                                            color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Indigo);
                                            break;
                                        case 12:
                                            color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Lime);
                                            break;
                                        case 13:
                                            color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Firebrick);
                                            break;
                                        case 14:
                                            color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Tomato);
                                            break;
                                        default:
                                            color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                                            break;
                                    }
                                    ColorMeMine.Characters[ind+1, index].Font.Color = color;
                                    ind = index;
                                    ptr = ptr + 1;
                                }*/
                            }
                            else
                            {
                                worksheet.Cells[i + 1, j + 1] = "";
                            }
                        }
                    }

                    //Getting the location and file name of the excel to save from user. 
                    SaveFileDialog saveDialog = new SaveFileDialog();
                    saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                    saveDialog.FilterIndex = 2;

                    if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        workbook.SaveAs(saveDialog.FileName);
                        MessageBox.Show("Export Successful", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                finally
                {
                    app.Quit();
                    workbook = null;
                    worksheet = null;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message.ToString()); }
        }




    }

    }

public static class StringHelper
{
    public static IEnumerable<int> findAll(this string str, string ch)
    {
        var index = 0;
        while (true)
        {
            index = str.IndexOf(ch, index + 1);
            if (index == -1)
                break;
            yield return index;
        }
    }
}
