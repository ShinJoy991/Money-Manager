using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Data.Common;

namespace app1
{
    public partial class Form1 : Form
    {
        // Get the directory of the executable file
        static string exeDirectory = AppDomain.CurrentDomain.BaseDirectory;

        // Define the relative paths to the files
        static string directoryPath = Path.Combine(exeDirectory, "Directory");
        string workbookPath = Path.Combine(directoryPath, "tien.xlsx");
        string workbookPath2 = Path.Combine(directoryPath, "tien2.xlsx");
        string current_date_path = Path.Combine(directoryPath, "date.txt");
        string dateSelectedPath = Path.Combine(directoryPath, "dateselected.txt");
        public Form1()
        {
            InitializeComponent();

            // Check if the files exist
            if (!File.Exists(workbookPath) || !File.Exists(workbookPath2) || !File.Exists(current_date_path))
            {
                // Show a warning message
                DialogResult result = MessageBox.Show("One or more Directory files are missing.", "Warning");
                // Cancel the application closing
                Environment.Exit(0);
            }
        }

        string gt = null;
        string s = "";
        int nutcapnhat = 1;
        int hang = 2;
        int cot = 2;

        private void Form1_Load(object sender, EventArgs e) // When open app
        {
            string datetime2string = File.ReadAllText(dateSelectedPath);
            Thread.Sleep(20);
            try
            {
                DateTime datetime2 = DateTime.ParseExact(datetime2string, "MM/dd/yy", null); // Convert String to DateTime var
                dateTimePicker2.Value = datetime2;
            }
            catch (Exception)
            {
                DateTime datetime2 = new DateTime(2000, 1, 1);
                dateTimePicker2.Value = datetime2;
            }
        }
        private string getExcel(Excel.Worksheet exSheet, int a, int b)
        {
            try
            {
                return Convert.ToString(exSheet.Cells[a, b].Value2);
            }
            catch (Exception e)
            {
                return "error";
            }
        }
        private int getExcelNum(Excel.Worksheet exSheet, int a, int b)
        {
            try
            {
                return int.Parse(Convert.ToString(exSheet.Cells[a, b].Value2));
            }
            catch (Exception e)
            {
                return 0;
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = new Excel.Application();

            Excel.Workbook exBook1 = exApp.Workbooks.Open(workbookPath,
                    0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);

            exApp.Visible = true;
            // get sheet 1.
            Excel.Worksheet exSheet = (Excel.Worksheet)exBook1.Worksheets["10DIEM"];

            System.Runtime.InteropServices.Marshal.ReleaseComObject(exBook1);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(exApp);
        }


        private void Xóa_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            label5.Text = "";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }


        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            DateTime selectedDate = dateTimePicker2.Value;
            Thread.Sleep(10);
            File.WriteAllText(dateSelectedPath, selectedDate.ToString("MM/dd/yy"));
            if(selectedDate > dateTimePicker3.Value)
            {
                textBox3.Text = "Wrong date";
            }
            else
            {
                ChenhLech();
            }
        }
        private void ChenhLech()
        {
            //Open excel
            Excel.Application exApp = new Excel.Application();
            Excel.Workbook exBook1 = exApp.Workbooks.Open(workbookPath,
                    0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);
            exApp.Visible = false;
            Excel.Worksheet exSheet = (Excel.Worksheet)exBook1.Worksheets["10DIEM"];
            // Calculate date
            DateTime datechenhlech1 = dateTimePicker2.Value;
            DateTime datechenhlech2 = dateTimePicker3.Value;
            int hangdau = 2;
            int hangsau = 2;
            int hangcuoi = exSheet.Cells[exSheet.Rows.Count, 1].End[XlDirection.xlUp].Row;
            if ( hangcuoi == 1)
            {
                goto ChenhLech2;
            }
            if (nutcapnhat == 1)
            {
                string old_date = File.ReadAllText(current_date_path);
                if (old_date != "'" + datechenhlech2.ToString("MM/dd/yy"))
                {
                    hangsau = hangcuoi;
                    goto ChenhLech1;
                }
            }
            for (int a = hangcuoi; a > 0; a--)
            {
                if (getExcel(exSheet,a,1) == datechenhlech2.ToString("MM/dd/yy"))
                {
                    hangsau = a;
                    break;
                }
            }
            ChenhLech1:
            for (int b = hangsau; b >= 0; b--)
            {
                if (b == 0)
                {
                    textBox3.Text = "Check start date!";
                    goto ChenhLech2;
                }
                if (getExcel(exSheet,b,1) == datechenhlech1.ToString("MM/dd/yy"))
                {
                    hangdau = b;
                    break;
                }
            }
            Thread.Sleep(10);
            int sum = 0;
            // Iterate through the rows in the specified range and calculate the sum
            for (int row = hangdau; row <= hangsau; row++)
            {
                // Get the value of the cell in the specified column and row
                int cellValue = getExcelNum(exSheet, row, 11);

                // Add the cell value to the sum
                sum += cellValue;
            }
            textBox3.Text = Convert.ToString(sum);
        ChenhLech2:
            exBook1.Save();
            exBook1.Close();
            exApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(exBook1);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(exApp);
        }
        private void button7_Click(object sender, EventArgs e)
        {
            if (nutcapnhat == 1)
            {
                ExcelFileReader(workbookPath);
                ExcelFileReader2();
                ChenhLech();
                nutcapnhat = nutcapnhat + 1;
            }
            if (textBox1.Text == "" || textBox2.Text == "")
            {
                label5.Text = "No data available!";
                label5.ForeColor = Color.Red;
                goto last;
            }
            Excel.Application exApp = new Excel.Application();

            Excel.Workbook exBook1 = exApp.Workbooks.Open(workbookPath,
                    0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);
            exApp.Visible = false;
            Excel.Worksheet exSheet = (Excel.Worksheet)exBook1.Worksheets["10DIEM"];


            DateTime new_date = dateTimePicker1.Value;
    //        int hang = 4;
            string str = "";
            hang = exSheet.Cells[exSheet.Rows.Count, 1].End[XlDirection.xlUp].Row;
    //        textBox3.Text = Convert.ToString(hang);
            // row is date and collumn is date assign box
            string old_date = File.ReadAllText(current_date_path);

            if (old_date == "'" + new_date.ToString("MM/dd/yy"))
            {
                Thread.Sleep(10);
            }
            else
            {
                hang = hang + 2;
                File.WriteAllText(current_date_path, String.Empty);
                Thread.Sleep(20);
                File.WriteAllText(current_date_path, "'" + new_date.ToString("MM/dd/yy"));
                Thread.Sleep(20);
            if( hang == 3)
                {
                    hang = 2;
                }
                Excel.Range date1 = (Excel.Range)exSheet.Cells[hang, 1];
                date1.Value2 = "'" + new_date.ToString("MM/dd/yy");
                Excel.Range tong = (Excel.Range)exSheet.Cells[hang, 11];
                tong.Value2 = "=SUM(B" + hang + ":J" + hang + ")";
              //  textBox3.Text = Convert.ToString(12);
                //               exSheet.Cells[hang, 1].Columns.AutoFit();
                //               exSheet.Cells[hang, 1].Rows.AutoFit();
            }
            for (cot = 2; cot <= 11; cot++)
            {
                str = getExcel(exSheet,hang,cot);

                if (str == null  && getExcel(exSheet,hang + 1,cot) == null)
                {
                    break;
                }
            }
            if (cot >= 11)
            {
                label5.Text = "";
                Thread.Sleep(100);
                label5.Text = "Use too much!! :(";
                label5.ForeColor = Color.Black;
                exBook1.Save();
                exBook1.Close();
                exApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(exBook1);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(exApp);
                goto last;
            }
            // Excel2
            CheckTen();
            //        Console.WriteLine(hang);
            //        Console.WriteLine(cot);
            Excel.Range tien = (Excel.Range)exSheet.Cells[hang, cot];
            tien.Value2 = textBox1.Text;
            Excel.Range lydo = (Excel.Range)exSheet.Cells[hang + 1, cot];
            lydo.Value2 = textBox2.Text;
            exSheet.Cells[hang, 12] = exSheet.Cells[5, 13].Value2;

            DateTime datechenhlech1 = dateTimePicker2.Value;
            DateTime datechenhlech2 = dateTimePicker3.Value;
            int hangdau = 2;
            int hangsau = 2;
            int hangcuoi = exSheet.Cells[exSheet.Rows.Count, 1].End[XlDirection.xlUp].Row;
            if (hangcuoi == 1)
            {
                goto ChenhLech2;
            }
            for (int a = hangcuoi; a > 0; a--)
            {
                if (getExcel(exSheet, a, 1) == datechenhlech2.ToString("MM/dd/yy"))
                {
                    hangsau = a;
                    break;
                }
            }
        ChenhLech1:
            for (int b = hangsau; b >= 0; b--)
            {
                if (b == 0)
                {
                    textBox3.Text = "Check start date!";
                    goto ChenhLech2;
                }
                if (getExcel(exSheet, b, 1) == datechenhlech1.ToString("MM/dd/yy"))
                {
                    hangdau = b;
                    break;
                }
            }
            Thread.Sleep(10);
            int sum = 0;
            // Iterate through the rows in the specified range and calculate the sum
            for (int row = hangdau; row <= hangsau; row++)
            {
                // Get the value of the cell in the specified column and row
                int cellValue = getExcelNum(exSheet, row, 11);

                // Add the cell value to the sum
                sum += cellValue;
            }
            textBox3.Text = Convert.ToString(sum);
        ChenhLech2:

            // Get data range
            Range range = exSheet.UsedRange;
            // Read data from sheet to 2D array
            object[,] values = range.Value;
            // Data table from 2D array
            var dataTable = new System.Data.DataTable();
            for (int col = 1; col <= values.GetLength(1); col++)
            {
                dataTable.Columns.Add(values[1, col]?.ToString() ?? "");
            }
            for (int row = 2; row <= values.GetLength(0); row++)
            {
                var dataRow = dataTable.NewRow();
                for (int col = 1; col <= values.GetLength(1); col++)
                {
                    dataRow[col - 1] = values[row, col]?.ToString() ?? "";
                }
                dataTable.Rows.Add(dataRow);
            }
            // Show on DataGridView
            dataGridView1.DataSource = dataTable;
            Thread.Sleep(10);
            dataGridView1.CurrentCell = dataGridView1.Rows[hang ].Cells[cot];
            ExcelFileReader2();

            exBook1.Save();
            exBook1.Close();
            exApp.Quit();
            Thread.Sleep(10);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(exBook1);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(exApp);
            label5.Text = "";
            Thread.Sleep(70);
            label5.Text = "Money updated";
            label5.ForeColor = Color.Green;
            goto lastbt7;
        last:;
            ExcelFileReader(workbookPath);
            ExcelFileReader2();
            label5.Text = "Updated sheet!";
            label5.ForeColor = Color.Green;
        lastbt7:;
        }
 
        public void ExcelFileReader(string path)
        {
               Excel.Application exApp = new Excel.Application();
               Excel.Workbook exBook1 = exApp.Workbooks.Open(workbookPath,
                       0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                       true, false, 0, true, false, false);
               // Get sheet 1.
               Excel.Worksheet exSheet = (Excel.Worksheet)exBook1.Worksheets["10DIEM"];
               // Get data range
               Range range = exSheet.UsedRange;
            // Read data from sheet to 2D array
            object[,] values = range.Value;
            // Data table from 2D array
            var dataTable = new System.Data.DataTable();
            for (int col = 1; col <= values.GetLength(1); col++)
               {
                   dataTable.Columns.Add(values[1, col]?.ToString() ?? "");
               }
               for (int row = 2; row <= values.GetLength(0); row++)
               {
                   var dataRow = dataTable.NewRow();
                   for (int col = 1; col <= values.GetLength(1); col++)
                   {
                       dataRow[col - 1] = values[row, col]?.ToString() ?? "";
                   }
                   dataTable.Rows.Add(dataRow);
               }
            // Show on DataGridView
            dataGridView1.DataSource = dataTable;
            Thread.Sleep(10);
            dataGridView1.CurrentCell = dataGridView1.Rows[hang - 2].Cells[cot];
            // Close Excel
            exBook1.Close();
            exApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(exBook1);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(exApp);
        }
        public void ExcelFileReader2()
        {
            Excel.Application exApp2 = new Excel.Application();
            Excel.Workbook exBook2 = exApp2.Workbooks.Open(workbookPath2,
                    0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);
            // Get sheet 1.
            Excel.Worksheet exSheet2 = (Excel.Worksheet)exBook2.Worksheets["10DIEM"];
            // Get the data range
            Range range = exSheet2.UsedRange;
            // Read data from the sheet into a 2-dimensional array
            object[,] values = range.Value;
            // Create a data table from the 2-dimensional array
            var dataTable = new System.Data.DataTable();
            for (int col = 1; col <= values.GetLength(1); col++)
            {
                dataTable.Columns.Add(values[1, col]?.ToString() ?? "");
            }
            for (int row = 1; row <= values.GetLength(0); row++)
            {
                var dataRow = dataTable.NewRow();
                for (int col = 1; col <= values.GetLength(1); col++)
                {
                    dataRow[col - 1] = values[row, col]?.ToString() ?? "";
                }
                dataTable.Rows.Add(dataRow);
            }
            // Display the data table on the DataGridView
            dataGridView2.DataSource = dataTable;
            // Close the Excel file
            exBook2.Save();
            exBook2.Close();
            exApp2.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(exBook2);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(exApp2);
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("Are you want to quit?" , "Warning" ,MessageBoxButtons.OKCancel) != System.Windows.Forms.DialogResult.OK)
            {
                e.Cancel = true;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            // Check if there are characters that are not numbers or "-" at the beginning of the string
            if (!System.Text.RegularExpressions.Regex.IsMatch(textBox1.Text, "^-?[0-9]*$"))
            {
                label5.Text = "Please enter a positive or negative number";
                label5.ForeColor = Color.Black;
                textBox1.Text = textBox1.Text.Remove(textBox1.SelectionStart - 1, 1); // Remove the last character
                textBox1.SelectionStart = textBox1.Text.Length; // Set the text cursor at the end of the TextBox
            }

            if (textBox1.Text.StartsWith("-"))
            {
                if (textBox1.Text.Length > 9)
                {
                    label5.Text = "Please enter a maximum of 8 digits";
                    label5.ForeColor = Color.Black;
                    textBox1.Text = textBox1.Text.Remove(textBox1.SelectionStart - 1, 1); // Remove the character just entered
                    textBox1.SelectionStart = textBox1.Text.Length; // Set the text cursor at the end of the TextBox
                }
                goto endtextbox1;
            }
            if (textBox1.Text.Length > 8)
            {
                label5.Text = "Please enter up to 8 digits";
                label5.ForeColor = Color.Black;
                textBox1.Text = textBox1.Text.Remove(textBox1.SelectionStart - 1, 1); // Remove the character just entered
                textBox1.SelectionStart = textBox1.Text.Length; // Set the text cursor at the end of the TextBox
            }
        endtextbox1:
            ;
        }



        // Design section

        private void tienchenhlech()
        {
            for (DateTime date = dateTimePicker2.Value; date <= dateTimePicker3.Value; date = date.AddDays(1))
            {

                Console.WriteLine(date.ToString("dd/MM/yyyy"));
            }
        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePicker3.Value < dateTimePicker2.Value)
            {
                textBox3.Text = "Wrong date!";
            }
            else
            {
                ChenhLech();
            }
        }

        private void CheckTen()
        {
            Excel.Application exApp2 = new Excel.Application();
            Excel.Workbook exBook2 = exApp2.Workbooks.Open(workbookPath2,
                    0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);
            // Get sheet 1.
            Excel.Worksheet exSheet2 = (Excel.Worksheet)exBook2.Worksheets["10DIEM"];
            string name1 = "name1";
            string name2 = "name2";
            string name3 = "name3";
            string name4 = "name4";
            string name5 = "name5";
            try { name1 = getExcel(exSheet2, 2, 2); } catch { }
            try { name2 = getExcel(exSheet2, 2, 3); } catch { }
            try { name3 = getExcel(exSheet2, 2, 4); } catch { }
            try { name4 = getExcel(exSheet2, 2, 5); } catch { }
            try { name5 = getExcel(exSheet2, 2, 6); } catch { }

            if (textBox2.Text.StartsWith(name1 ?? "name1"))
            {
                // Handle if the text starts with "name1"
                Excel.Range range = exSheet2.Cells[3, 2];
                // Get the current value of the cell and add 3
                int currentValue = 0;
                try
                {
                    currentValue = (int)range.Value;
                }
                catch (Exception e) {
                    currentValue = 0;
                }
                range.Value = currentValue + int.Parse(textBox1.Text);
            }
            if (textBox2.Text.StartsWith(name2 ?? "name2"))
            {
                // Handle if the text starts with "Tan"
                Excel.Range range = exSheet2.Cells[3, 3];
                int currentValue = 0;
                try
                {
                    currentValue = (int)range.Value;
                }
                catch (Exception e)
                {
                    currentValue = 0;
                }
                // Get the current value of the cell and add
                range.Value = currentValue + int.Parse(textBox1.Text);
            }
            if (textBox2.Text.StartsWith(name3 ?? "name3"))
            {
                Excel.Range range = exSheet2.Cells[3, 5];
                int currentValue = 0;
                try
                {
                    currentValue = (int)range.Value;
                }
                catch (Exception e)
                {
                    currentValue = 0;
                }
                range.Value = currentValue + int.Parse(textBox1.Text);
            }
            if (textBox2.Text.StartsWith(name4 ?? "name4"))
            {
                Excel.Range range = exSheet2.Cells[3, 4];
                int currentValue = 0;
                try
                {
                    currentValue = (int)range.Value;
                }
                catch (Exception e)
                {
                    currentValue = 0;
                }
                range.Value = currentValue + int.Parse(textBox1.Text);
            }
            if (textBox2.Text.StartsWith(name5 ?? "name5"))
            {
                Excel.Range range = exSheet2.Cells[3, 6];
                int currentValue = 0;
                try
                {
                    currentValue = (int)range.Value;
                }
                catch (Exception e)
                {
                    currentValue = 0;
                }
                range.Value = currentValue + int.Parse(textBox1.Text);
            }
            exBook2.Save();
            exBook2.Close();
            exApp2.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(exBook2);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(exApp2);
        }

    }

}

