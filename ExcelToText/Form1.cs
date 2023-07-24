using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string cellValue = null;
            bool stop = false;
            StringBuilder line = new StringBuilder();
            string textLine = "";
            do
            {
                string excelFilePath = string.Empty;
                openFileDialog1.InitialDirectory = Application.StartupPath;
                openFileDialog1.FileName = "*.xlsx";
                openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    excelFilePath = openFileDialog1.FileName;
                    Excel.Application xlApp = new Excel.Application();
                    //Excel.Workbook xlWorkbook = xlApp.Workbooks.Add();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(excelFilePath);
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;
                    int rowCount = xlRange.Rows.Count;
                    int colCount = xlRange.Columns.Count;
                    string[] list = new string[colCount];
                    //List<string[]> lists = new List<string[]> ();
                    string[][] lists = new string[rowCount][];
                    for (int i = 0; i <= rowCount - 1; i++)
                    {
                        lists[i] = list;
                        for (int j = 0; j <= colCount - 1; j++)
                        {
                            if (xlRange.Cells[i + 1, j + 1] != null && xlRange.Cells[i + 1, j + 1].Value2 != null)
                            {
                                cellValue = xlRange.Cells[i + 1, j + 1].Value2;
                                list[j] = cellValue.ToString();
                                label1.Text += string.Format("{0,-7}", cellValue.ToString());
                            }
                        }
                        lists[i] = list;
                        label1.Text += "\r\n";
                    }
                    xlApp.Workbooks.Close();
                    xlApp.Quit();

                    xlRange = null;
                    xlWorksheet = null;
                    xlWorkbook = null;
                    xlApp = null;
                    stop = true;
                    for (int i = 0; i <= rowCount - 1; i++)
                    {
                        list = lists[i];
                        for (int j = 0; j <= colCount - 1; j++)
                        {
                            textLine += string.Format("{0,-20}", list[j]);
                        }
                        textLine += "\r\n";
                    }
                    string textFilePath = string.Empty;
                    saveFileDialog1.InitialDirectory = Application.StartupPath;
                    saveFileDialog1.FileName = "*.txt";
                    saveFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        textFilePath = saveFileDialog1.FileName;
                        StreamWriter sw = new StreamWriter(textFilePath);
                        for (int i = 0; i <= rowCount - 1; i++)
                        {
                            list = lists[i];
                            for (int j = 0; j <= colCount - 1; j++)
                            {
                                sw.WriteLine(string.Format("{0,10}", list[j]));
                            }
                            sw.WriteLine("\r\n");
                        }
                        sw.Close();
                    }
                    else
                    {
                        MessageBox.Show("Dosya kaydedilemedi!", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    if (label1.Text == " ")
                    {
                        MessageBox.Show("Excel dosyasını seçmediniz!", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        this.button1_Click(sender, e);
                        break;
                    }

                }
            }
            while (stop == false);
        }

    }
}
