using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Media;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace Inventory
{
    public partial class Form1 : Form
    {

        public string[][] LoadList(string filePath)
        {


            StreamReader sr = new StreamReader(@filePath);
            var lines = new List<string[]>();
            int Row = 0;
            while (!sr.EndOfStream)
            {
                string[] Line = sr.ReadLine().Split('\t');
                lines.Add(Line);
                Row++;
                //Console.WriteLine(Row);
            }
            return lines.ToArray();
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e) // LOAD FILE
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = Directory.GetCurrentDirectory();
            openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx";
            //openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                
                System.IO.FileInfo fInfo = new System.IO.FileInfo(openFileDialog1.FileName);

                string strFileName = fInfo.Name;
                string strFilePath = fInfo.DirectoryName;
                string path = strFilePath + "\\" + strFileName;
               
                try
                {
                    var workbook = new XLWorkbook(path);
                    var worksheet = workbook.Worksheet(1);
                    int new_input = 0;
                    int last_row = worksheet.LastRowUsed().RowNumber();

                    if ((string)worksheet.Cell("A1").Value == "DESCRIPTION")
                    {
                        new_input = 1;
                    }

                    for (int i = 1 + new_input; i < last_row; i++)
                    {
                        // Cell "A" is Description, Cell "B" is asset, Cell "C" is check.
                        bool check = false;
                        var column_c = worksheet.Cell("C" + i).Value;
                        if (Object.ReferenceEquals(column_c.GetType(), "bool"))
                        {
                            check = true;
                        }

                        if ((string)worksheet.Cell("A" + i).Value != "")
                        {
                            dataGridView1.Rows.Add((string)worksheet.Cell("A" + i).Value, worksheet.Cell("B" + i).Value, check);
                        }
                    }

                    dataGridView1.Refresh();

                    return;
                }
                catch (System.IO.IOException)
                {
                    MessageBox.Show("File open in another window. Please close and try again.");
                    return;
                }
                

                
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e) // DATA ENTRY
        {

        }
        private void playSimpleSound()
        {
            SoundPlayer simpleSound = new SoundPlayer(@"c:\Windows\Media\chimes.wav");
            simpleSound.Play();
        }
        private void playSimpleSound2()
        {
            SoundPlayer simpleSound = new SoundPlayer(@"c:\Windows\Media\chord.wav");
            simpleSound.Play();
        }

        private void button2_Click(object sender, EventArgs e) // ENTER DATA BUTTON
        {
            string look = "";
            string search = textBox1.Text;
            if (search != "")
            {
                //MessageBox.Show(search);
                bool found = false;
                int size = dataGridView1.RowCount;
                int i = 0;
                //MessageBox.Show(size.ToString());
                while (!found && i < size)
                {
                    look = dataGridView1.Rows[i].Cells[1].Value.ToString();

                    if (look == search)
                    {
                        found = true;
                        dataGridView1.Rows[i].Cells[2].Value = true;
                    }
                    i++;
                }
                if (!found)
                {
                    playSimpleSound2();
                    MessageBox.Show("Not found");
                }
                else
                { playSimpleSound(); }
                textBox1.Text = "";
            }
        }

        private void button3_Click(object sender, EventArgs e) // SAVE FILE
        {
            saveFile(false);
        }

        private void button4_Click_1(object sender, EventArgs e) // "PRINT" (meaning save, then open in windows)
        {
            saveFile(true);
        }

        private void saveFile(bool openFile)
        {
            // Displays a SaveFileDialog so the user can save the file
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "XLSX File|*.xlsx";
            saveFileDialog1.Title = "Save your inventory";
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                // Saves the file via a FileStream created by the OpenFile method.  
                System.IO.FileInfo fInfo = new System.IO.FileInfo(saveFileDialog1.FileName);

                string path = fInfo.DirectoryName + "\\" + fInfo.Name;

                var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Serial Inventory");
                int row_size = dataGridView1.RowCount;
                for (int i = 0; i < row_size - 1; i++)
                {
                    // there are only 3 columns: description, asset, and true/false (which is simply to see if it is checked)
                    worksheet.Cell("A" + (i + 1)).Value = dataGridView1.Rows[i].Cells[0].Value;
                    worksheet.Cell("B" + (i + 1)).Value = dataGridView1.Rows[i].Cells[1].Value;
                    worksheet.Cell("C" + (i + 1)).Value = dataGridView1.Rows[i].Cells[2].Value;
                }
                worksheet.Columns().AdjustToContents();

                path = Regex.Replace(path, ".xlsx", "");

                workbook.SaveAs(path + ".xlsx"); // this does NOT correctly interpret filenames that DO contain ".xlsx"
                if (openFile)
                {
                    Process process = new Process();
                    Process.Start("Excel.exe", path);
                }
            }
        }

        private void printPreviewDialog1_Load(object sender, EventArgs e)
        {
            ShowDialog();
        }

        public string[,] GetData()
        {
            string[,] Data = { };
            int j = 0;
            int size = dataGridView1.RowCount;
            for (int i = 0; i < size - 1; i++)
            {
                if (dataGridView1.Rows[i].Cells[2].Value != null && !(bool)dataGridView1.Rows[i].Cells[2].Value)
                {
                    Data[j, 0] = (string)dataGridView1.Rows[i].Cells[0].Value;
                    Data[j, 1] = (string)dataGridView1.Rows[i].Cells[1].Value;
                }
            }
            return Data;
        }

        

        // Figure out how wide each column should be.
        private int[] FindColumnWidths(Graphics gr, Font header_font,
            Font body_font, string[] headers, string[,] values)
        {
            // Make room for the widths.
            int[] widths = new int[headers.Length];

            // Find the width for each column.
            for (int col = 0; col < widths.Length; col++)
            {
                // Check the column header.
                widths[col] = (int)gr.MeasureString(
                    headers[col], header_font).Width;

                // Check the items.
                for (int row = 0; row <= values.GetUpperBound(0); row++)
                {
                    int value_width = (int)gr.MeasureString(
                        values[row, col], body_font).Width;
                    if (widths[col] < value_width)
                        widths[col] = value_width;
                }

                // Add some extra space.
                widths[col] += 20;
            }

            return widths;
        }
    }
}
