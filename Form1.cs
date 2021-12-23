using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            // Open the Excel file using ClosedXML.
            // Keep in mind the Excel file cannot be open when trying to read it
            using (XLWorkbook workBook = new XLWorkbook(filePath))
            {
                //Read the first Sheet from Excel file.
                IXLWorksheet workSheet = workBook.Worksheet(1);

                //Create a new DataTable.
                DataTable dt = new DataTable();

                //Loop through the Worksheet rows.
                bool firstRow = true;
                foreach (IXLRow row in workSheet.Rows())
                {
                    //Use the first row to add columns to DataTable.
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
                        //Add rows to DataTable.
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
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
            {
                filePath = file.FileName; //get the path of the file  
                fp = Path.GetDirectoryName(filePath);
                fileExt = Path.GetExtension(filePath); //get the file extension  
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        DataTable dtExcel = new DataTable();
                        dtExcel = ReadExcel(filePath); //read excel file  
                        dataGridView1.Visible = true;
                        dataGridView1.DataSource = dtExcel;
                        textBox1.Text = filePath;
                        label5.Text = fp;
                        DataRowCollection rows1 = dtExcel.Rows;
                        DataColumnCollection cols1 = dtExcel.Columns;
                        textBox3.Text = rows1.Count.ToString();
                        textBox4.Text = Convert.ToString(cols1.Count);

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
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            Parse(Convert.ToInt32(textBox3.Text), Convert.ToInt32(textBox5.Text), 
                ReadExcel(textBox1.Text), label5.Text);
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

        private void Parse(int rowN, int fileN, DataTable dtExcel, String fp)
        {
            int i = 1;
            IXLWorkbook wb1 = new XLWorkbook(textBox1.Text);
            IXLWorksheet ws1 = wb1.Worksheet(1);

            IXLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");

            ws1.Cell(1, 1).CopyTo(ws.Cell(1, 1));
            wb.SaveAs(fp+ "Parsed_"+i+".xlsx");


        }
    }
}
