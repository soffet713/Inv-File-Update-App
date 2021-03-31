using System;
using System.IO;
using System.Threading;
using System.Collections.Generic;
using OpenQA.Selenium;
using System.Windows.Forms;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

/************************************************************************************************
**
**  Date: March 31th, 2021
**  Application Name: Inv File Update Application
**  Author: Sean McWilliams
**
**  Description: Application that takes Investment Excel file with worksheet named Performance
**               Summary and updates each percentage for different Benchmark Funds.
**
**  Current File: Application main Form code.
**
***********************************************************************************************/

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

        //Dictionary with links to pull current investment performance summary percentages
        IDictionary<int, string> psum_sitelinks = new Dictionary<int, string>()
        {
            { 0, "https://fundresearch.fidelity.com/mutual-funds/performance-and-risk/31635V729" },
            { 1, "https://fundresearch.fidelity.com/mutual-funds/performance-and-risk/31635T815" },
            { 2, "https://fundresearch.fidelity.com/mutual-funds/performance-and-risk/31635T781" },
            { 3, "https://markets.ft.com/data/etfs/tearsheet/performance?s=IDEV:PCQ:USD" },
            { 4, "https://markets.ft.com/data/etfs/tearsheet/performance?s=GEM:PCQ:USD" },
            { 5, "https://www.ishares.com/us/products/239623/ishares-msci-eafe-etf" },
            { 6, "https://www.spglobal.com/spdji/en/indices/fixed-income/sp-us-treasury-bill-index/#overview" },
            { 7, "https://www.spglobal.com/spdji/en/indices/fixed-income/sp-500-bond-index/#overview" }
        };

        //Dictionary with links to pull current investment bond performance percentages
        IDictionary<int, string> bperf_sitelinks = new Dictionary<int, string>()
        {
            { 0, "https://www.spglobal.com/spdji/en/indices/fixed-income/sp-taxable-municipal-bond-index/#overview" },
            { 1, "https://www.spglobal.com/spdji/en/indices/fixed-income/sp-us-government-bond-index/#overview" },
            { 2, "https://www.morningstar.com/funds/xnas/crefx/performance" },
            { 3, "https://fundresearch.fidelity.com/mutual-funds/performance-and-risk/74925K367" },
            { 4, "https://www.ishares.com/us/products/239561/ishares-gold-trust-fund" },
            { 5, "https://fundresearch.fidelity.com/mutual-funds/performance-and-risk/00203H602" },
            { 6, "https://www.blackrock.com/us/individual/products/227660/blackrock-strategic-income-opportunitiesinst-cl-fund" },
            { 7, "https://fundresearch.fidelity.com/mutual-funds/performance-and-risk/64128R608" },
            { 8, "https://www.ishares.com/us/products/239855/ishares-silver-trust-fund" },
            { 9, "https://www.ssga.com/us/en/institutional/etfs/funds/spdr-gold-shares-gld" },
            { 10, "https://www.spglobal.com/spdji/en/indices/equity/sp-500-real-estate-index/#overview" },
            { 11, "https://markets.ft.com/data/funds/tearsheet/performance?s=TILIX" },
            { 12, "https://markets.ft.com/data/funds/tearsheet/performance?s=TILVX" },
            { 13, "https://markets.ft.com/data/funds/tearsheet/performance?s=MFOMX" },
            { 14, "https://markets.ft.com/data/funds/tearsheet/performance?s=MPMCX" },
            { 15, "https://markets.ft.com/data/funds/tearsheet/performance?s=DMVYX" },
            { 16, "https://markets.ft.com/data/funds/tearsheet/performance?s=DSGYX" },
            { 17, "https://markets.ft.com/data/funds/tearsheet/performance?s=DISYX" },
            { 18, "https://markets.ft.com/data/funds/tearsheet/performance?s=NIEYX" },
            { 19, "https://markets.ft.com/data/funds/tearsheet/performance?s=MEMKX" },
            { 20, "https://markets.ft.com/data/funds/tearsheet/performance?s=TIHYX" }
        };

        string sFileName;

        private void Form1_Load(object sender, EventArgs e)
        {
            if (comboBox1.Text.Equals(null) || comboBox1.Text == "")
            {
                updateButton.Enabled = false;
            }
            saveFileButton.Enabled = false;
        }

        //Function call on Update Button click that calls two functions:
        private void updateButton_Click(object sender, EventArgs e)
        {
            //Functions calls of Pull_perf_sum_data() and perf_cell_update to get current percentages from web and sends values to cells in Excel file
            if (comboBox1.Text == "Performance Summary")
            {
                FunctionsPage my_page = new FunctionsPage();
                IJavaScriptExecutor js = my_page as IJavaScriptExecutor;
                string[] values;
                my_page.Manage().Window.Maximize();
                for (int i = 0; i < psum_sitelinks.Count; i++)
                {
                    my_page.Navigate().GoToUrl(psum_sitelinks[i]);
                    Thread.Sleep(TimeSpan.FromSeconds(3));
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
                        case 7:
                            values = my_page.Pull_perf_sum_data(i);
                            perf_cell_update(null, 30, values);
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

            //Functions calls of Pull_bond_perf_data() and bond_perf_cell_update to get current percentages from web and sends values to cells in Excel file
            if (comboBox1.Text == "Bond Performance")
            {
                FunctionsPage bp_page = new FunctionsPage();
                IJavaScriptExecutor js = bp_page as IJavaScriptExecutor;
                string[] bp_values;
                bp_page.Manage().Window.Maximize();
                for (int p = 0; p < bperf_sitelinks.Count; p++)
                {
                    bp_page.Navigate().GoToUrl(bperf_sitelinks[p]);
                    Thread.Sleep(TimeSpan.FromSeconds(3));
                    switch (p)
                    {
                        case 0:
                            bp_values = bp_page.Pull_bond_perf_data(p);
                            bond_perf_cell_update(4, 9, bp_values);
                            break;
                        case 1:
                            bp_values = bp_page.Pull_bond_perf_data(p);
                            bond_perf_cell_update(5, 9, bp_values);
                            break;
                        case 2:
                            bp_values = bp_page.Pull_bond_perf_data(p);
                            bond_perf_cell_update(18, 3, bp_values);
                            break;
                        case 3:
                            bp_values = bp_page.Pull_bond_perf_data(p);
                            bond_perf_cell_update(19, 3, bp_values);
                            break;
                        case 4:
                            bp_values = bp_page.Pull_bond_perf_data(p);
                            bond_perf_cell_update(20, 3, bp_values);
                            break;
                        case 5:
                            bp_values = bp_page.Pull_bond_perf_data(p);
                            bond_perf_cell_update(21, 3, bp_values);
                            break;
                        case 6:
                            bp_values = bp_page.Pull_bond_perf_data(p);
                            bond_perf_cell_update(22, 3, bp_values);
                            break;
                        case 7:
                            bp_values = bp_page.Pull_bond_perf_data(p);
                            bond_perf_cell_update(23, 3, bp_values);
                            break;
                        case 8:
                            bp_values = bp_page.Pull_bond_perf_data(p);
                            bond_perf_cell_update(24, 3, bp_values);
                            break;
                        case 9:
                            bp_values = bp_page.Pull_bond_perf_data(p);
                            bond_perf_cell_update(25, 3, bp_values);
                            break;
                        case 10:
                            bp_values = bp_page.Pull_bond_perf_data(p);
                            bond_perf_cell_update(14, 13, bp_values);
                            break;
                        case 11:
                            bp_values = bp_page.Pull_bond_perf_data(p);
                            bond_perf_cell_update(41, 3, bp_values);
                            break;
                        case 12:
                            bp_values = bp_page.Pull_bond_perf_data(p);
                            bond_perf_cell_update(42, 3, bp_values);
                            break;
                        case 13:
                            bp_values = bp_page.Pull_bond_perf_data(p);
                            bond_perf_cell_update(43, 3, bp_values);
                            break;
                        case 14:
                            bp_values = bp_page.Pull_bond_perf_data(p);
                            bond_perf_cell_update(44, 3, bp_values);
                            break;
                        case 15:
                            bp_values = bp_page.Pull_bond_perf_data(p);
                            bond_perf_cell_update(45, 3, bp_values);
                            break;
                        case 16:
                            bp_values = bp_page.Pull_bond_perf_data(p);
                            bond_perf_cell_update(46, 3, bp_values);
                            break;
                        case 17:
                            bp_values = bp_page.Pull_bond_perf_data(p);
                            bond_perf_cell_update(47, 3, bp_values);
                            break;
                        case 18:
                            bp_values = bp_page.Pull_bond_perf_data(p);
                            bond_perf_cell_update(48, 3, bp_values);
                            break;
                        case 19:
                            bp_values = bp_page.Pull_bond_perf_data(p);
                            bond_perf_cell_update(49, 3, bp_values);
                            break;
                        case 20:
                            bp_values = bp_page.Pull_bond_perf_data(p);
                            bond_perf_cell_update(50, 3, bp_values);
                            break;
                    }
                }
                bp_page.Close();
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
                //Traverse through each percentage value, check if string already has percentage sign attached, and send it to correct cell of Excel worksheet
                string num1 = vals[x];
                if (!num1.EndsWith("%"))
                {
                    num1 += "%";
                }
                
                if (x <= 4 && row1.HasValue)
                {
                    xlWorkSheet.Cells[row1, x + 3].Value = num1;
                }
                else if (x > 4 && row2.HasValue)
                {
                    xlWorkSheet.Cells[row2, x - 2].Value = num1;
                }
            }
        }

        private void bond_perf_cell_update(int row, int col, string[] bp_vals)
        {
            for (int y = 0; y < bp_vals.Length; y++)
            {
                string num2 = bp_vals[y];
                if (!num2.EndsWith("%"))
                {
                    num2 += "%";
                }

                xlWorkSheet.Cells[row, y + col].Value = num2;
            }
        }

        //Allows user to open Investment file and loads Performance Summary worksheet into dropdown box
        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Excel File to Edit";
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Excel File|*.xlsx;*.xls";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                sFileName = openFileDialog1.FileName;
                int filepathlength = sFileName.Length - sFileName.IndexOf(".");
                string newFileName = sFileName.Substring(0, (sFileName.Length - filepathlength)) + " (Copy)" + sFileName.Substring(sFileName.IndexOf("."), filepathlength);
                File.Copy(sFileName, newFileName);
                sFileName = newFileName;

                if (sFileName.Trim() != "")
                {
                    textBox1.Text = sFileName;
                    xlWorkBook = xlApp.Workbooks.Open(sFileName);
                    foreach (Excel.Worksheet ws in xlWorkBook.Worksheets)
                    {
                        if (ws.Name == "Performance Summary" || ws.Name == "Bond Performance")
                        {
                            comboBox1.Items.Add(ws.Name);
                        }
                    }
                }
            }
        }

        //Function to allow user to save file either with a new file name or replace original file with updated percentage information
        private void button2_Click(object sender, EventArgs e)
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

        //Checks value of dropdown box and changes status of the Update button
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
