using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace BulkFileParser
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public static DataTable ReadExcel(string filePath)
        {
         
            using (XLWorkbook workBook = new XLWorkbook(filePath))
            {
               
                IXLWorksheet workSheet = workBook.Worksheet(1);

                
                DataTable dt = new DataTable();

               
                bool firstRow = true;
                foreach (IXLRow row in workSheet.Rows())
                {
                   
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Columns.Add(cell.Value.ToString());
                        }
                        firstRow = false;
                    }
                    else
                    {
                        dt.Rows.Add();
                        int i = 0;

                        foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                            i++;
                        }
                    }
                }

                return dt;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            string fp = string.Empty;
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) 
            {
                filePath = file.FileName; 
                fp = Path.GetDirectoryName(filePath);
                fileExt = Path.GetExtension(filePath); 
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        DataTable dtExcel = new DataTable();
                        dtExcel = ReadExcel(filePath);
                        dataGridView1.Visible = true;
                        dataGridView1.DataSource = dtExcel;
                        textBox1.Text = filePath;
                        label5.Text = fp;
                        DataRowCollection rows1 = dtExcel.Rows;
                        DataColumnCollection cols1 = dtExcel.Columns;
                        textBox3.Text = rows1.Count.ToString();
                        textBox4.Text = cols1.Count.ToString();

                        if (textBox2.Text == "")
                        {
                            textBox5.Text = "Input row cut off number";
                        }
                        else
                            textBox2_TextChanged(sender,  e);
                            textBox5.Enabled = true;



                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);  
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            progressBar1.Visible = true;
            progressBar1.Maximum = Convert.ToInt32(textBox5.Text);
            Parse(Convert.ToInt32(textBox2.Text), Convert.ToInt32(textBox4.Text), Convert.ToInt32(textBox5.Text), 
                ReadExcel(textBox1.Text), label5.Text);
            progressBar1.Visible = false;

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (textBox3.Text == "")
            {
                textBox5.Enabled = false;
            }
            else
            {
                int rpf = Convert.ToInt32(textBox2.Text);
                textBox5.Text = Convert.ToString((Convert.ToInt32(textBox3.Text) / rpf)+1);
            }
            
        }

        private void Parse(int rowN, int colN, int fileN, DataTable dtExcel, String fp)
        {
            
            IXLWorkbook wb1 = new XLWorkbook(textBox1.Text);
            IXLWorksheet ws1 = wb1.Worksheet(1);
            IXLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Clear((XLClearOptions)11);
            int n;     
            int lRow;

            if (checkBox1.Checked)
            {
                lRow = rowN + 1;
                n = 2;
            }
            else
            {
                lRow = rowN;
                n = 1;
            }
            int fRow = n;

            for (int i = 1; i <= fileN; i++)
            {
                if (checkBox1.Checked)
                    ws1.Range(ws1.Cell(1, 1), ws1.Cell(1, colN)).CopyTo(ws.Range("A1"));
                   
                
                ws1.Range(ws1.Cell(fRow, 1), ws1.Cell(lRow, colN)).CopyTo(ws.Range("A"+n));
                
                    
                wb.SaveAs(fp + "\\Parsed_" + i + ".xlsx");
                fRow = lRow + 1;
                lRow = rowN+ lRow;
                progressBar1.Increment(1);

            }
            
        }
        

        
    }
}
