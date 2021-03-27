using System;
using System.Threading;
using System.Collections.Generic;
using OpenQA.Selenium;
using System.Windows.Forms;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace Inv_File_Update_App
{
    
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;

        IDictionary<int, string> sitelinks = new Dictionary<int, string>()
        {
            { 0, "https://fundresearch.fidelity.com/mutual-funds/performance-and-risk/31635V729" },
            { 1, "https://fundresearch.fidelity.com/mutual-funds/performance-and-risk/31635T815" },
            { 2, "https://fundresearch.fidelity.com/mutual-funds/performance-and-risk/31635T781" },
            { 3, "https://markets.ft.com/data/etfs/tearsheet/performance?s=IDEV:PCQ:USD" },
            { 4, "https://markets.ft.com/data/etfs/tearsheet/performance?s=GEM:PCQ:USD" },
            { 5, "https://www.ishares.com/us/products/239623/ishares-msci-eafe-etf" },
            { 6, "https://www.spglobal.com/spdji/en/indices/fixed-income/sp-us-treasury-bill-index/#overview" }
        };

        string sFileName;

        private void Form1_Load(object sender, EventArgs e)
        {
            if (comboBox1.Text.Equals(null) || comboBox1.Text == "")
            {
                updateButton.Enabled = false;
            }
            saveFileButton.Enabled = false;
            //closeFileButton.Enabled = false;
        }

        private void updateButton_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "Performance Summary")
            {
                
                FunctionsPage my_page = new FunctionsPage();
                IJavaScriptExecutor js = my_page as IJavaScriptExecutor;
                string[] values;
                my_page.Manage().Window.Maximize();
                for (int i = 0; i < sitelinks.Count; i++)
                {
                    my_page.Navigate().GoToUrl(sitelinks[i]);
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    switch (i)
                    {
                        case 0:
                            values = my_page.Pull_perf_sum_data(i);
                            perf_cell_update(8, 9, values);
                            break;
                        case 1:
                            values = my_page.Pull_perf_sum_data(i);
                            perf_cell_update(11, 12, values);
                            break;
                        case 2:
                            values = my_page.Pull_perf_sum_data(i);
                            perf_cell_update(14, 15, values);
                            break;
                        case 3:
                            values = my_page.Pull_perf_sum_data(i);
                            perf_cell_update(17, 18, values);
                            break;
                        case 4:
                            values = my_page.Pull_perf_sum_data(i);
                            perf_cell_update(20, 21, values);
                            break;
                        case 5:
                            values = my_page.Pull_perf_sum_data(i);
                            perf_cell_update(23, 24, values);
                            break;
                        case 6:
                            values = my_page.Pull_perf_sum_data(i);
                            perf_cell_update(26, 27, values);
                            break;
                    }
                    xlWorkSheet.Cells[3, 1].Value = "Performance Summary";
                }
                my_page.Close();
                Process[] chromeDriverProcesses = Process.GetProcessesByName("chromedriver");
                foreach (var chromeDriverProcess in chromeDriverProcesses)
                {
                    chromeDriverProcess.Kill();
                }
            }
            saveFileButton.Enabled = true;
        }

        private void perf_cell_update(int? row1, int? row2, string[] vals)
        {
            for (int x = 0; x < vals.Length; x++)
            {
                string num = vals[x];
                if (!num.EndsWith("%"))
                {
                    num += "%";
                }
                
                if (x <= 4 && row1.HasValue)
                {
                    xlWorkSheet.Cells[row1, x + 3].Value = num;
                }
                else if (row2.HasValue)
                {
                    xlWorkSheet.Cells[row2, x - 2].Value = num;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Excel File to Edit";
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Excel File|*.xlsx;*.xls";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                sFileName = openFileDialog1.FileName;

                if (sFileName.Trim() != "")
                {
                    textBox1.Text = sFileName;
                    xlWorkBook = xlApp.Workbooks.Open(sFileName);
                    foreach (Excel.Worksheet ws in xlWorkBook.Worksheets)
                    {
                        if (ws.Name == "Performance Summary")
                        {
                            comboBox1.Items.Add(ws.Name);
                        }
                        Console.WriteLine(ws.Name);
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Title = "Excel File to Edit";
            saveFileDialog1.FileName = "";
            saveFileDialog1.Filter = "Excel File|*.xlsx;*.xls";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                sFileName = saveFileDialog1.FileName;

                xlWorkBook.SaveAs(sFileName);
            }
            //closeFileButton.Enabled = true;
            button3_Click(sender, e);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                xlWorkBook.Close();
                xlApp.Quit();
                if (xlWorkSheet != null) { releaseObject(xlWorkSheet); };
                if (xlWorkBook != null) { releaseObject(xlWorkBook); };
                releaseObject(xlApp);
                foreach (Process clsProcess in Process.GetProcesses())
                {
                    if (clsProcess.ProcessName.Equals("EXCEL"))
                    {
                        clsProcess.Kill();
                        break;
                    }
                }
            }
            catch
            {
                Exception ex = new Exception("File is already closed.");
                MessageBox.Show(ex.ToString());
            }
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (comboBox1.Text.Equals(null) || comboBox1.Text == "") { updateButton.Enabled = false; }
                else
                {
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets[comboBox1.Text];
                    updateButton.Enabled = true;
                }
            }
            catch
            {
                throw new Exception("Error updating combo box with worksheet names.");
            }
        }
    }
}
