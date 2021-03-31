using System;
using System.Collections.Generic;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Collections.ObjectModel;

/************************************************************************************************
**
**  Date: March 31th, 2021
**  Application Name: Inv File Update Application
**  Author: Sean McWilliams
**
**  Description: Application that takes Investment Excel file with worksheet named Performance
**               Summary and updates each percentage for different Benchmark Funds.
**
**  Current File: Selenium function file to pull most recent percentages from the web.
**
***********************************************************************************************/

namespace Inv_File_Update_App
{
    class FunctionsPage : ChromeDriver
    {
        //Selenium function to pull data from browser for each item in top Dictionary of links for Performance Summary worksheet
        //Returns string array of new percentages
        public String[] Pull_perf_sum_data(int l)
        {
            string path1_mtd, path1_3mo, path1_ytd, path1_1yr, path1_3yr, path2_mtd, path2_3mo, path2_ytd, path2_1yr, path2_3yr;
            string value1_mtd = "", value1_3mo = "", value1_1yr = "", value1_ytd = "", value1_3yr = "", value2_mtd = "", value2_3mo = "", value2_1yr = "", value2_ytd = "", value2_3yr = "";
            
            switch (l)
            {
                case 0:
                    path1_mtd = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::span[contains(text(),'Large Growth')]" +
                        "/ancestor::tr/child::td[3]/child::span";
                    path1_3mo = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::span[contains(text(),'Large Growth')]" +
                        "/ancestor::tr/child::td[4]/child::span";
                    path1_ytd = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::span[contains(text(),'Large Growth')]" +
                        "/ancestor::tr/child::td[2]/child::span";
                    path1_1yr = "//span[@class='ng-binding'][contains(text(),'Large Growth')]/parent::td/following-sibling::td[@class='avg-annual--table-cell']/" +
                        "child::span[@class='ng-binding ng-scope']";
                    path1_3yr = "//span[@class='ng-binding'][contains(text(),'Large Growth')]/parent::td/following-sibling::td[@class='avg-annual--table-cell'][2]" +
                        "/child::span[@class='ng-binding ng-scope']";
                    path2_mtd = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::a[contains(text(),'Russell 1000 Growth')]" +
                        "/ancestor::tr/child::td[3]/child::span";
                    path2_3mo = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::a[contains(text(),'Russell 1000 Growth')]" +
                        "/ancestor::tr/child::td[4]/child::span";
                    path2_ytd = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::a[contains(text(),'Russell 1000 Growth')]" +
                        "/ancestor::tr/child::td[2]/child::span";
                    path2_1yr = "//a[@class='glossary-link ng-binding ng-scope'][contains(text(),'Russell 1000 Growth')]/parent::td/following-sibling::td[@class='avg-annual--table-cell']" +
                        "/child::span[@class='ng-binding ng-scope']";
                    path2_3yr = "//a[@class='glossary-link ng-binding ng-scope'][contains(text(),'Russell 1000 Growth')]/parent::td/following-sibling::td[@class='avg-annual--table-cell'][2]" +
                        "/child::span[@class='ng-binding ng-scope']";
                    value1_mtd = FindElement(By.XPath(path1_mtd)).Text;
                    value1_3mo = FindElement(By.XPath(path1_3mo)).Text;
                    value1_ytd = FindElement(By.XPath(path1_ytd)).Text;
                    value1_1yr = FindElement(By.XPath(path1_1yr)).Text;
                    value1_3yr = FindElement(By.XPath(path1_3yr)).Text;
                    value2_mtd = FindElement(By.XPath(path2_mtd)).Text;
                    value2_3mo = FindElement(By.XPath(path2_3mo)).Text;
                    value2_ytd = FindElement(By.XPath(path2_ytd)).Text;
                    value2_1yr = FindElement(By.XPath(path2_1yr)).Text;
                    value2_3yr = FindElement(By.XPath(path2_3yr)).Text;
                    break;
                case 1:
                    path1_mtd = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::a[contains(text(),'Russell Midcap Growth')]" +
                        "/ancestor::tr/child::td[3]/child::span";
                    path1_3mo = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::a[contains(text(),'Russell Midcap Growth')]" +
                        "/ancestor::tr/child::td[4]/child::span";
                    path1_ytd = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::a[contains(text(),'Russell Midcap Growth')]" +
                        "/ancestor::tr/child::td[2]/child::span";
                    path1_1yr = "//a[@class='glossary-link ng-binding ng-scope'][contains(text(),'Russell Midcap Growth')]/parent::td/following-sibling::td[@class='avg-annual--table-cell']" +
                        "/child::span[@class='ng-binding ng-scope']";
                    path1_3yr = "//a[@class='glossary-link ng-binding ng-scope'][contains(text(),'Russell Midcap Growth')]/parent::td/following-sibling::td[@class='avg-annual--table-cell'][2]" +
                        "/child::span[@class='ng-binding ng-scope']";
                    path2_mtd = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::span[contains(text(),'Mid-Cap Growth')]" +
                        "/ancestor::tr/child::td[3]/child::span";
                    path2_3mo = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::span[contains(text(),'Mid-Cap Growth')]" +
                        "/ancestor::tr/child::td[4]/child::span";
                    path2_ytd = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::span[contains(text(),'Mid-Cap Growth')]" +
                        "/ancestor::tr/child::td[2]/child::span";
                    path2_1yr = "//span[@class='ng-binding'][contains(text(),'Mid-Cap Growth')]/parent::td/following-sibling::td[@class='avg-annual--table-cell']/" +
                        "child::span[@class='ng-binding ng-scope']";
                    path2_3yr = "//span[@class='ng-binding'][contains(text(),'Mid-Cap Growth')]/parent::td/following-sibling::td[@class='avg-annual--table-cell'][2]/" +
                        "child::span[@class='ng-binding ng-scope']";
                    value1_mtd = FindElement(By.XPath(path1_mtd)).Text;
                    value1_3mo = FindElement(By.XPath(path1_3mo)).Text;
                    value1_ytd = FindElement(By.XPath(path1_ytd)).Text;
                    value1_1yr = FindElement(By.XPath(path1_1yr)).Text;
                    value1_3yr = FindElement(By.XPath(path1_3yr)).Text;
                    value2_mtd = FindElement(By.XPath(path2_mtd)).Text;
                    value2_3mo = FindElement(By.XPath(path2_3mo)).Text;
                    value2_ytd = FindElement(By.XPath(path2_ytd)).Text;
                    value2_1yr = FindElement(By.XPath(path2_1yr)).Text;
                    value2_3yr = FindElement(By.XPath(path2_3yr)).Text;
                    break;
                case 2:
                    path1_mtd = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::span[contains(text(),'Small Growth')]" +
                        "/ancestor::tr/child::td[3]/child::span";
                    path1_3mo = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::span[contains(text(),'Small Growth')]" +
                        "/ancestor::tr/child::td[4]/child::span";
                    path1_ytd = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::span[contains(text(),'Small Growth')]" +
                        "/ancestor::tr/child::td[2]/child::span";
                    path1_1yr = "//span[@class='ng-binding'][contains(text(),'Small Growth')]/parent::td/following-sibling::td[@class='avg-annual--table-cell']/" +
                        "child::span[@class='ng-binding ng-scope']";
                    path1_3yr = "//span[@class='ng-binding'][contains(text(),'Small Growth')]/parent::td/following-sibling::td[@class='avg-annual--table-cell'][2]" +
                        "/child::span[@class='ng-binding ng-scope']";
                    path2_mtd = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::a[contains(text(),'Russell 2000 Growth')]" +
                        "/ancestor::tr/child::td[3]/child::span";
                    path2_3mo = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::a[contains(text(),'Russell 2000 Growth')]" +
                        "/ancestor::tr/child::td[4]/child::span";
                    path2_ytd = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::a[contains(text(),'Russell 2000 Growth')]" +
                        "/ancestor::tr/child::td[2]/child::span";
                    path2_1yr = "//a[@class='glossary-link ng-binding ng-scope'][contains(text(),'Russell 2000 Growth')]/parent::td/following-sibling::td[@class='avg-annual--table-cell']" +
                        "/child::span[@class='ng-binding ng-scope']";
                    path2_3yr = "//a[@class='glossary-link ng-binding ng-scope'][contains(text(),'Russell 2000 Growth')]/parent::td/following-sibling::td[@class='avg-annual--table-cell'][2]" +
                        "/child::span[@class='ng-binding ng-scope']";
                    value1_mtd = FindElement(By.XPath(path1_mtd)).Text;
                    value1_3mo = FindElement(By.XPath(path1_3mo)).Text;
                    value1_ytd = FindElement(By.XPath(path1_ytd)).Text;
                    value1_1yr = FindElement(By.XPath(path1_1yr)).Text;
                    value1_3yr = FindElement(By.XPath(path1_3yr)).Text;
                    value2_mtd = FindElement(By.XPath(path2_mtd)).Text;
                    value2_3mo = FindElement(By.XPath(path2_3mo)).Text;
                    value2_ytd = FindElement(By.XPath(path2_ytd)).Text;
                    value2_1yr = FindElement(By.XPath(path2_1yr)).Text;
                    value2_3yr = FindElement(By.XPath(path2_3yr)).Text;
                    break;
                case 3:
                    path1_mtd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'iShares Core MSCI International Developed Markets ETF')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    path1_3mo = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'iShares Core MSCI International Developed Markets ETF')]/parent::td" +
                        "/following-sibling::td[5]/child::span";
                    path1_1yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'iShares Core MSCI International Developed Markets ETF')]/parent::td" +
                        "/following-sibling::td[3]/child::span";
                    path1_3yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'iShares Core MSCI International Developed Markets ETF')]/parent::td" +
                        "/following-sibling::td[2]/child::span";
                    path2_mtd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'MSCI World Ex USA IMI NR USD')]/parent::td/following-sibling::td[6]" +
                        "/child::span";
                    path2_3mo = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'MSCI World Ex USA IMI NR USD')]/parent::td/following-sibling::td[5]" +
                        "/child::span";                    
                    path2_1yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'MSCI World Ex USA IMI NR USD')]/parent::td/following-sibling::td[3]" +
                        "/child::span";
                    path2_3yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'MSCI World Ex USA IMI NR USD')]/parent::td/following-sibling::td[2]" +
                        "/child::span";                    
                    value1_mtd = FindElement(By.XPath(path1_mtd)).Text.Trim(new Char[] { '+' });
                    value1_3mo = FindElement(By.XPath(path1_3mo)).Text.Trim(new Char[] { '+' });
                    value1_1yr = FindElement(By.XPath(path1_1yr)).Text.Trim(new Char[] { '+' });
                    value1_3yr = FindElement(By.XPath(path1_3yr)).Text.Trim(new Char[] { '+' });
                    value2_mtd = FindElement(By.XPath(path2_mtd)).Text.Trim(new Char[] { '+' });
                    value2_3mo = FindElement(By.XPath(path2_3mo)).Text.Trim(new Char[] { '+' });
                    value2_1yr = FindElement(By.XPath(path2_1yr)).Text.Trim(new Char[] { '+' });
                    value2_3yr = FindElement(By.XPath(path2_3yr)).Text.Trim(new Char[] { '+' });
                    FindElement(By.XPath("//li[@class='mod-ui-tab mod-ui-tab__module-header'][contains(text(),'Annual')]")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    path1_ytd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'iShares Core MSCI International Developed Markets ETF')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    path2_ytd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'MSCI World Ex USA IMI NR USD')]/parent::td/following-sibling::td[6]" +
                        "/child::span";
                    value1_ytd = FindElement(By.XPath(path1_ytd)).Text.Trim(new Char[] { '+' });
                    value2_ytd = FindElement(By.XPath(path2_ytd)).Text.Trim(new Char[] { '+' });
                    break;
                case 4:
                    path1_mtd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'Goldman Sachs ActiveBeta® Emerging Markets Equity ETF')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    path1_3mo = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'Goldman Sachs ActiveBeta® Emerging Markets Equity ETF')]/parent::td" +
                        "/following-sibling::td[5]/child::span";
                    path1_1yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'Goldman Sachs ActiveBeta® Emerging Markets Equity ETF')]/parent::td" +
                        "/following-sibling::td[3]/child::span";
                    path1_3yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'Goldman Sachs ActiveBeta® Emerging Markets Equity ETF')]/parent::td" +
                        "/following-sibling::td[2]/child::span";
                    path2_mtd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'Diversified Emerging Mkts')]/parent::td/following-sibling::td[6]" +
                        "/child::span";
                    path2_3mo = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'Diversified Emerging Mkts')]/parent::td/following-sibling::td[5]" +
                        "/child::span";
                    path2_1yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'Diversified Emerging Mkts')]/parent::td/following-sibling::td[3]" +
                        "/child::span";
                    path2_3yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'Diversified Emerging Mkts')]/parent::td/following-sibling::td[2]" +
                        "/child::span";                    
                    value1_mtd = FindElement(By.XPath(path1_mtd)).Text.Trim(new Char[] { '+' });
                    value1_3mo = FindElement(By.XPath(path1_3mo)).Text.Trim(new Char[] { '+' });
                    value1_1yr = FindElement(By.XPath(path1_1yr)).Text.Trim(new Char[] { '+' });
                    value1_3yr = FindElement(By.XPath(path1_3yr)).Text.Trim(new Char[] { '+' });
                    value2_mtd = FindElement(By.XPath(path2_mtd)).Text.Trim(new Char[] { '+' });
                    value2_3mo = FindElement(By.XPath(path2_3mo)).Text.Trim(new Char[] { '+' });
                    value2_1yr = FindElement(By.XPath(path2_1yr)).Text.Trim(new Char[] { '+' });
                    value2_3yr = FindElement(By.XPath(path2_3yr)).Text.Trim(new Char[] { '+' });
                    FindElement(By.XPath("//li[@class='mod-ui-tab mod-ui-tab__module-header'][contains(text(),'Annual')]")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    path1_ytd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'Goldman Sachs ActiveBeta® Emerging Markets Equity ETF')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    value1_ytd = FindElement(By.XPath(path1_ytd)).Text.Trim(new Char[] { '+' });
                    path2_ytd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'Diversified Emerging Mkts')]/parent::td/following-sibling::td[6]" +
                        "/child::span";
                    value2_ytd = FindElement(By.XPath(path2_ytd)).Text.Trim(new Char[] { '+' });
                    break;
                case 5:
                    FindElement(By.XPath("//a[@class='cumulative'][contains(text(), 'Cumulative')]")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    FindElement(By.XPath("//div[@id='subTabCumulative']/child::div[@id='cumulativeTabs']/child::div[@class='component-date-list']/descendant::select[@class='date-dropdown']")).Click();
                    FindElement(By.XPath("//div[@id='subTabCumulative']/child::div[@id='cumulativeTabs']/child::div[@class='component-date-list']/descendant::select[@class='date-dropdown']/child::option[1]")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    path1_mtd = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[1]/child::td[@class='oneMonth ']";
                    path1_3mo = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[1]/child::td[@class='threeMonth ']";
                    path1_ytd = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[1]/child::td[@class='yearToDate ']";
                    path1_1yr = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[1]/child::td[@class='oneYear ']";
                    path1_3yr = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[1]/child::td[@class='threeYear ']";
                    path2_mtd = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[3]/child::td[@class='oneMonth ']";
                    path2_3mo = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[3]/child::td[@class='threeMonth ']";
                    path2_ytd = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[3]/child::td[@class='yearToDate ']";
                    path2_1yr = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[3]/child::td[@class='oneYear ']";
                    path2_3yr = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[3]/child::td[@class='threeYear ']";
                    value1_mtd = FindElement(By.XPath(path1_mtd)).Text;
                    value1_3mo = FindElement(By.XPath(path1_3mo)).Text;
                    value1_ytd = FindElement(By.XPath(path1_ytd)).Text;
                    value1_1yr = FindElement(By.XPath(path1_1yr)).Text;
                    value1_3yr = FindElement(By.XPath(path1_3yr)).Text;
                    value2_mtd = FindElement(By.XPath(path2_mtd)).Text;
                    value2_3mo = FindElement(By.XPath(path2_3mo)).Text;
                    value2_ytd = FindElement(By.XPath(path2_ytd)).Text;
                    value2_1yr = FindElement(By.XPath(path2_1yr)).Text;
                    value2_3yr = FindElement(By.XPath(path2_3yr)).Text;
                    break;
                case 6:
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    FindElement(By.XPath("//button[@id='onetrust-accept-btn-handler']")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    FindElement(By.XPath("//div[@class='content-pane performance']/descendant::li[@class='chart active']/following-sibling::li[@class='table-view ']")).Click();
                    FindElement(By.XPath("//div[@class='content-pane performance']/descendant::div[@class='performance-right-table-panel']" +
                        "/descendant::li[@submoduleflag='MonthEnd']")).Click();
                    path1_mtd = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'1 MTH')]/following-sibling::span";
                    path1_3mo = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'3 MTH')]/following-sibling::span";
                    path1_ytd = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'YTD')]/following-sibling::span";
                    path1_1yr = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'1 Year')]/following-sibling::span";
                    path1_3yr = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'3 Year')]/following-sibling::span";
                    value1_mtd = FindElement(By.XPath(path1_mtd)).Text;
                    value1_3mo = FindElement(By.XPath(path1_3mo)).Text;
                    value1_ytd = FindElement(By.XPath(path1_ytd)).Text;
                    value1_1yr = FindElement(By.XPath(path1_1yr)).Text;
                    value1_3yr = FindElement(By.XPath(path1_3yr)).Text;
                    //Go to additional link for second set of percentages
                    Navigate().GoToUrl("https://www.spglobal.com/spdji/en/indices/fixed-income/sp-us-ultra-short-treasury-bill-bond-index/#overview");
                    Thread.Sleep(TimeSpan.FromSeconds(3));
                    FindElement(By.XPath("//div[@class='content-pane performance']/descendant::li[@class='chart active']/following-sibling::li[@class='table-view ']")).Click();
                    FindElement(By.XPath("//div[@class='content-pane performance']/descendant::div[@class='performance-right-table-panel']" +
                        "/descendant::li[@submoduleflag='MonthEnd']")).Click();
                    path2_mtd = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'1 MTH')]/following-sibling::span";
                    path2_3mo = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'3 MTH')]/following-sibling::span";
                    path2_ytd = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'YTD')]/following-sibling::span";
                    path2_1yr = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'1 Year')]/following-sibling::span";
                    path2_3yr = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'3 Year')]/following-sibling::span";
                    value2_mtd = FindElement(By.XPath(path2_mtd)).Text;
                    value2_3mo = FindElement(By.XPath(path2_3mo)).Text;
                    value2_ytd = FindElement(By.XPath(path2_ytd)).Text;
                    value2_1yr = FindElement(By.XPath(path2_1yr)).Text;
                    value2_3yr = FindElement(By.XPath(path2_3yr)).Text;
                    break;
                case 7:
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    FindElement(By.XPath("//div[@class='content-pane performance']/descendant::li[@class='chart active']/following-sibling::li[@class='table-view ']")).Click();
                    FindElement(By.XPath("//div[@class='content-pane performance']/descendant::div[@class='performance-right-table-panel']" +
                        "/descendant::li[@submoduleflag='MonthEnd']")).Click();
                    path2_mtd = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'1 MTH')]/following-sibling::span";
                    path2_3mo = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'3 MTH')]/following-sibling::span";
                    path2_ytd = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'YTD')]/following-sibling::span";
                    path2_1yr = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'1 Year')]/following-sibling::span";
                    path2_3yr = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'3 Year')]/following-sibling::span";
                    value2_mtd = FindElement(By.XPath(path2_mtd)).Text;
                    value2_3mo = FindElement(By.XPath(path2_3mo)).Text;
                    value2_ytd = FindElement(By.XPath(path2_ytd)).Text;
                    value2_1yr = FindElement(By.XPath(path2_1yr)).Text;
                    value2_3yr = FindElement(By.XPath(path2_3yr)).Text;
                    break;
            }
            string[] vals = { value1_mtd, value1_3mo, value1_ytd, value1_1yr, value1_3yr, value2_mtd, value2_3mo, value2_ytd, value2_1yr, value2_3yr };
            return vals;
        }

        //Selenium function to pull data from browser for each item in top Dictionary of links for Bond Performance worksheet
        //Returns string array of new percentages
        public String[] Pull_bond_perf_data(int p)
        {
            string path_mtd, path_3mo, path_ytd, path_1yr, path_3yr;
            string value_mtd = "", value_3mo = "", value_1yr = "", value_ytd = "", value_3yr = "";

            switch (p)
            {
                case 0:
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    FindElement(By.XPath("//button[@id='onetrust-accept-btn-handler']")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    FindElement(By.XPath("//div[@class='content-pane performance']/descendant::li[@class='chart active']/following-sibling::li[@class='table-view ']")).Click();
                    FindElement(By.XPath("//div[@class='content-pane performance']/descendant::div[@class='performance-right-table-panel']" +
                        "/descendant::li[@submoduleflag='MonthEnd']")).Click();
                    path_mtd = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'1 MTH')]/following-sibling::span";
                    path_3mo = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'3 MTH')]/following-sibling::span";
                    path_ytd = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'YTD')]/following-sibling::span";
                    path_1yr = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'1 Year')]/following-sibling::span";
                    path_3yr = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'3 Year')]/following-sibling::span";
                    value_mtd = FindElement(By.XPath(path_mtd)).Text;
                    value_3mo = FindElement(By.XPath(path_3mo)).Text;
                    value_ytd = FindElement(By.XPath(path_ytd)).Text;
                    value_1yr = FindElement(By.XPath(path_1yr)).Text;
                    value_3yr = FindElement(By.XPath(path_3yr)).Text;
                    break;
                case 1:
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    FindElement(By.XPath("//div[@class='content-pane performance']/descendant::li[@class='chart active']/following-sibling::li[@class='table-view ']")).Click();
                    FindElement(By.XPath("//div[@class='content-pane performance']/descendant::div[@class='performance-right-table-panel']" +
                        "/descendant::li[@submoduleflag='MonthEnd']")).Click();
                    path_mtd = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'1 MTH')]/following-sibling::span";
                    path_3mo = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'3 MTH')]/following-sibling::span";
                    path_ytd = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'YTD')]/following-sibling::span";
                    path_1yr = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'1 Year')]/following-sibling::span";
                    path_3yr = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'3 Year')]/following-sibling::span";
                    value_mtd = FindElement(By.XPath(path_mtd)).Text;
                    value_3mo = FindElement(By.XPath(path_3mo)).Text;
                    value_ytd = FindElement(By.XPath(path_ytd)).Text;
                    value_1yr = FindElement(By.XPath(path_1yr)).Text;
                    value_3yr = FindElement(By.XPath(path_3yr)).Text;
                    break;
                case 2:
                    FindElement(By.XPath("//section[@class='sal-performance-trailing-return']/descendant::div[@class='sal-trailing-return__filters']/descendant::div[@class='mwc-tabs']" +
                        "/descendant::mds-button[@value='Month End']")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    path_mtd = "//section[@class='sal-performance-trailing-return']/descendant::div[@class='sal-component-body']/descendant::table[@class='total-table']/child::tbody" +
                        "/child::tr[@class='total-return']/child::td[1]";
                    path_3mo = "//section[@class='sal-performance-trailing-return']/descendant::div[@class='sal-component-body']/descendant::table[@class='total-table']/child::tbody" +
                        "/child::tr[@class='total-return']/child::td[2]";
                    path_ytd = "//section[@class='sal-performance-trailing-return']/descendant::div[@class='sal-component-body']/descendant::table[@class='total-table']/child::tbody" +
                        "/child::tr[@class='total-return']/child::td[3]";
                    path_1yr = "//section[@class='sal-performance-trailing-return']/descendant::div[@class='sal-component-body']/descendant::table[@class='total-table']/child::tbody" +
                        "/child::tr[@class='total-return']/child::td[4]";
                    path_3yr = "//section[@class='sal-performance-trailing-return']/descendant::div[@class='sal-component-body']/descendant::table[@class='total-table']/child::tbody" +
                        "/child::tr[@class='total-return']/child::td[5]";
                    value_mtd = FindElement(By.XPath(path_mtd)).Text;
                    value_3mo = FindElement(By.XPath(path_3mo)).Text;
                    value_ytd = FindElement(By.XPath(path_ytd)).Text;
                    value_1yr = FindElement(By.XPath(path_1yr)).Text;
                    value_3yr = FindElement(By.XPath(path_3yr)).Text;
                    break;
                case 3:
                    path_mtd = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::span[contains(text(),'Abbey Capital Futures Strategy Fund Class I')]" +
                        "/ancestor::tr/child::td[3]/child::span";
                    path_3mo = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::span[contains(text(),'Abbey Capital Futures Strategy Fund Class I')]" +
                        "/ancestor::tr/child::td[4]/child::span";
                    path_ytd = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::span[contains(text(),'Abbey Capital Futures Strategy Fund Class I')]" +
                        "/ancestor::tr/child::td[2]/child::span";
                    path_1yr = "//span[@class='ng-binding'][contains(text(),'Abbey Capital Futures Strategy Fund Class I')]/parent::td/following-sibling::td[@class='avg-annual--table-cell']/" +
                        "child::span[@class='ng-binding ng-scope']";
                    path_3yr = "//span[@class='ng-binding'][contains(text(),'Abbey Capital Futures Strategy Fund Class I')]/parent::td/following-sibling::td[@class='avg-annual--table-cell'][2]" +
                        "/child::span[@class='ng-binding ng-scope']";
                    value_mtd = FindElement(By.XPath(path_mtd)).Text;
                    value_3mo = FindElement(By.XPath(path_3mo)).Text;
                    value_ytd = FindElement(By.XPath(path_ytd)).Text;
                    value_1yr = FindElement(By.XPath(path_1yr)).Text;
                    value_3yr = FindElement(By.XPath(path_3yr)).Text;
                    break;
                case 4:
                    FindElement(By.XPath("//a[@class='cumulative'][contains(text(), 'Cumulative')]")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    FindElement(By.XPath("//div[@id='subTabCumulative']/child::div[@id='cumulativeTabs']/child::div[@class='component-date-list']/descendant::select[@class='date-dropdown']")).Click();
                    FindElement(By.XPath("//div[@id='subTabCumulative']/child::div[@id='cumulativeTabs']/child::div[@class='component-date-list']/descendant::select[@class='date-dropdown']/child::option[1]")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    path_mtd = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[1]/child::td[@class='oneMonth ']";
                    path_3mo = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[1]/child::td[@class='threeMonth ']";
                    path_ytd = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[1]/child::td[@class='yearToDate ']";
                    path_1yr = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[1]/child::td[@class='oneYear ']";
                    path_3yr = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[1]/child::td[@class='threeYear ']";
                    value_mtd = FindElement(By.XPath(path_mtd)).Text;
                    value_3mo = FindElement(By.XPath(path_3mo)).Text;
                    value_ytd = FindElement(By.XPath(path_ytd)).Text;
                    value_1yr = FindElement(By.XPath(path_1yr)).Text;
                    value_3yr = FindElement(By.XPath(path_3yr)).Text;
                    break;
                case 5:
                    path_mtd = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::span[contains(text(),'AQR Diversified Arbitrage Fund Class I')]" +
                        "/ancestor::tr/child::td[3]/child::span";
                    path_3mo = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::span[contains(text(),'AQR Diversified Arbitrage Fund Class I')]" +
                        "/ancestor::tr/child::td[4]/child::span";
                    path_ytd = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::span[contains(text(),'AQR Diversified Arbitrage Fund Class I')]" +
                        "/ancestor::tr/child::td[2]/child::span";
                    path_1yr = "//span[@class='ng-binding'][contains(text(),'AQR Diversified Arbitrage Fund Class I')]/parent::td/following-sibling::td[@class='avg-annual--table-cell']/" +
                        "child::span[@class='ng-binding ng-scope']";
                    path_3yr = "//span[@class='ng-binding'][contains(text(),'AQR Diversified Arbitrage Fund Class I')]/parent::td/following-sibling::td[@class='avg-annual--table-cell'][2]" +
                        "/child::span[@class='ng-binding ng-scope']";
                    value_mtd = FindElement(By.XPath(path_mtd)).Text;
                    value_3mo = FindElement(By.XPath(path_3mo)).Text;
                    value_ytd = FindElement(By.XPath(path_ytd)).Text;
                    value_1yr = FindElement(By.XPath(path_1yr)).Text;
                    value_3yr = FindElement(By.XPath(path_3yr)).Text;
                    break;
                case 6:
                    FindElement(By.XPath("//a[@class='cumulative'][contains(text(), 'Cumulative')]")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    FindElement(By.XPath("//div[@id='subTabCumulative']/child::div[@id='cumulativeTabs']/child::div[@class='component-date-list']/descendant::select[@class='date-dropdown']")).Click();
                    FindElement(By.XPath("//div[@id='subTabCumulative']/child::div[@id='cumulativeTabs']/child::div[@class='component-date-list']/descendant::select[@class='date-dropdown']/child::option[1]")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    path_mtd = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[1]/child::td[@class='oneMonth ']";
                    path_3mo = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[1]/child::td[@class='threeMonth ']";
                    path_ytd = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[1]/child::td[@class='yearToDate ']";
                    path_1yr = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[1]/child::td[@class='oneYear ']";
                    path_3yr = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[1]/child::td[@class='threeYear ']";
                    value_mtd = FindElement(By.XPath(path_mtd)).Text;
                    value_3mo = FindElement(By.XPath(path_3mo)).Text;
                    value_ytd = FindElement(By.XPath(path_ytd)).Text;
                    value_1yr = FindElement(By.XPath(path_1yr)).Text;
                    value_3yr = FindElement(By.XPath(path_3yr)).Text;
                    break;
                case 7:
                    path_mtd = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::span[contains(text(),'Neuberger Berman Long Short Fund Institutional Class')]" +
                        "/ancestor::tr/child::td[3]/child::span";
                    path_3mo = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::span[contains(text(),'Neuberger Berman Long Short Fund Institutional Class')]" +
                        "/ancestor::tr/child::td[4]/child::span";
                    path_ytd = "//div[@id='cumulative-ret--table-div']/child::table[@class='cumulative-ret--table']/descendant::span[contains(text(),'Neuberger Berman Long Short Fund Institutional Class')]" +
                        "/ancestor::tr/child::td[2]/child::span";
                    path_1yr = "//span[@class='ng-binding'][contains(text(),'Neuberger Berman Long Short Fund Institutional Class')]/parent::td/following-sibling::td[@class='avg-annual--table-cell']/" +
                        "child::span[@class='ng-binding ng-scope']";
                    path_3yr = "//span[@class='ng-binding'][contains(text(),'Neuberger Berman Long Short Fund Institutional Class')]/parent::td/following-sibling::td[@class='avg-annual--table-cell'][2]" +
                        "/child::span[@class='ng-binding ng-scope']";
                    value_mtd = FindElement(By.XPath(path_mtd)).Text;
                    value_3mo = FindElement(By.XPath(path_3mo)).Text;
                    value_ytd = FindElement(By.XPath(path_ytd)).Text;
                    value_1yr = FindElement(By.XPath(path_1yr)).Text;
                    value_3yr = FindElement(By.XPath(path_3yr)).Text;
                    break;
                case 8:
                    FindElement(By.XPath("//a[@class='cumulative'][contains(text(), 'Cumulative')]")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    FindElement(By.XPath("//div[@id='subTabCumulative']/child::div[@id='cumulativeTabs']/child::div[@class='component-date-list']/descendant::select[@class='date-dropdown']")).Click();
                    FindElement(By.XPath("//div[@id='subTabCumulative']/child::div[@id='cumulativeTabs']/child::div[@class='component-date-list']/descendant::select[@class='date-dropdown']/child::option[1]")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    path_mtd = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[1]/child::td[@class='oneMonth ']";
                    path_3mo = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[1]/child::td[@class='threeMonth ']";
                    path_ytd = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[1]/child::td[@class='yearToDate ']";
                    path_1yr = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[1]/child::td[@class='oneYear ']";
                    path_3yr = "//table[@class='product-table border-row cumulative-returns']/child::tbody/child::tr[1]/child::td[@class='threeYear ']";
                    value_mtd = FindElement(By.XPath(path_mtd)).Text;
                    value_3mo = FindElement(By.XPath(path_3mo)).Text;
                    value_ytd = FindElement(By.XPath(path_ytd)).Text;
                    value_1yr = FindElement(By.XPath(path_1yr)).Text;
                    value_3yr = FindElement(By.XPath(path_3yr)).Text;
                    break;
                case 9:
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    FindElementById("js-ssmp-clrButtonLabel").Click();
                    Thread.Sleep(TimeSpan.FromSeconds(1));
                    FindElement(By.XPath("//a[@data-tab-name='#performance'][contains(text(),'Performance')]")).Click();
                    string activetab = FindElement(By.XPath("//div[@id='performance']/descendant::div[@class='table-chart']/descendant::div[@class='end tab-nav']" +
                        "/child::a[@class='tab-nav-link active']")).Text;
                    if(activetab != "Month End")
                    {
                        FindElement(By.XPath("//div[@id='performance']/descendant::div[@class='table-chart']/descendant::div[@class='end tab-nav']" +
                        "/child::a[contains(text(),'Month End')]")).Click();
                    }
                    string perftblpath = "//div[@id='performance']/descendant::div[@class='table-chart']/descendant::table[@class='tab-panel tab-end-ann active']/child::tbody/child::tr[2]/";
                    path_mtd = perftblpath + "child::td[3]";
                    path_3mo = perftblpath + "child::td[4]";
                    path_ytd = perftblpath + "child::td[5]";
                    path_1yr = perftblpath + "child::td[6]";
                    path_3yr = perftblpath + "child::td[7]";
                    value_mtd = FindElement(By.XPath(path_mtd)).Text;
                    value_3mo = FindElement(By.XPath(path_3mo)).Text;
                    value_ytd = FindElement(By.XPath(path_ytd)).Text;
                    value_1yr = FindElement(By.XPath(path_1yr)).Text;
                    value_3yr = FindElement(By.XPath(path_3yr)).Text;
                    break;
                case 10:
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    FindElement(By.XPath("//div[@class='content-pane performance']/descendant::li[@class='chart active']/following-sibling::li[@class='table-view ']")).Click();
                    FindElement(By.XPath("//div[@class='content-pane performance']/descendant::div[@class='performance-right-table-panel']" +
                        "/descendant::li[@submoduleflag='MonthEnd']")).Click();
                    path_mtd = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'1 MTH')]/following-sibling::span";
                    path_3mo = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'3 MTH')]/following-sibling::span";
                    path_ytd = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'YTD')]/following-sibling::span";
                    path_1yr = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'1 Year')]/following-sibling::span";
                    path_3yr = "//div[@class='performance-right-table-panel']/descendant::div[@class='annualized-return-table']/child::div[@class='data-row']" +
                        "/descendant::span[contains(text(),'3 Year')]/following-sibling::span";
                    value_mtd = FindElement(By.XPath(path_mtd)).Text;
                    value_3mo = FindElement(By.XPath(path_3mo)).Text;
                    value_ytd = FindElement(By.XPath(path_ytd)).Text;
                    value_1yr = FindElement(By.XPath(path_1yr)).Text;
                    value_3yr = FindElement(By.XPath(path_3yr)).Text;
                    break;
                case 11:
                    path_mtd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'TIAA-CREF Large-Cap Growth Index Fund Institutional Class')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    path_3mo = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'TIAA-CREF Large-Cap Growth Index Fund Institutional Class')]/parent::td" +
                        "/following-sibling::td[5]/child::span";
                    path_1yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'TIAA-CREF Large-Cap Growth Index Fund Institutional Class')]/parent::td" +
                        "/following-sibling::td[3]/child::span";
                    path_3yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'TIAA-CREF Large-Cap Growth Index Fund Institutional Class')]/parent::td" +
                        "/following-sibling::td[2]/child::span";
                    value_mtd = FindElement(By.XPath(path_mtd)).Text.Trim(new Char[] { '+' });
                    value_3mo = FindElement(By.XPath(path_3mo)).Text.Trim(new Char[] { '+' });
                    value_1yr = FindElement(By.XPath(path_1yr)).Text.Trim(new Char[] { '+' });
                    value_3yr = FindElement(By.XPath(path_3yr)).Text.Trim(new Char[] { '+' });
                    FindElement(By.XPath("//li[@class='mod-ui-tab mod-ui-tab__module-header'][contains(text(),'Annual')]")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    path_ytd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'TIAA-CREF Large-Cap Growth Index Fund Institutional Class')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    value_ytd = FindElement(By.XPath(path_ytd)).Text.Trim(new Char[] { '+' });
                    break;
                case 12:
                    path_mtd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'TIAA-CREF Large-Cap Value Index Fund Institutional Class')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    path_3mo = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'TIAA-CREF Large-Cap Value Index Fund Institutional Class')]/parent::td" +
                        "/following-sibling::td[5]/child::span";
                    path_1yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'TIAA-CREF Large-Cap Value Index Fund Institutional Class')]/parent::td" +
                        "/following-sibling::td[3]/child::span";
                    path_3yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'TIAA-CREF Large-Cap Value Index Fund Institutional Class')]/parent::td" +
                        "/following-sibling::td[2]/child::span";
                    value_mtd = FindElement(By.XPath(path_mtd)).Text.Trim(new Char[] { '+' });
                    value_3mo = FindElement(By.XPath(path_3mo)).Text.Trim(new Char[] { '+' });
                    value_1yr = FindElement(By.XPath(path_1yr)).Text.Trim(new Char[] { '+' });
                    value_3yr = FindElement(By.XPath(path_3yr)).Text.Trim(new Char[] { '+' });
                    FindElement(By.XPath("//li[@class='mod-ui-tab mod-ui-tab__module-header'][contains(text(),'Annual')]")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    path_ytd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'TIAA-CREF Large-Cap Value Index Fund Institutional Class')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    value_ytd = FindElement(By.XPath(path_ytd)).Text.Trim(new Char[] { '+' });
                    break;
                case 13:
                    path_mtd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Focused Equity Opportunities Fund Class M Shares')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    path_3mo = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Focused Equity Opportunities Fund Class M Shares')]/parent::td" +
                        "/following-sibling::td[5]/child::span";
                    path_1yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Focused Equity Opportunities Fund Class M Shares')]/parent::td" +
                        "/following-sibling::td[3]/child::span";
                    path_3yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Focused Equity Opportunities Fund Class M Shares')]/parent::td" +
                        "/following-sibling::td[2]/child::span";
                    value_mtd = FindElement(By.XPath(path_mtd)).Text.Trim(new Char[] { '+' });
                    value_3mo = FindElement(By.XPath(path_3mo)).Text.Trim(new Char[] { '+' });
                    value_1yr = FindElement(By.XPath(path_1yr)).Text.Trim(new Char[] { '+' });
                    value_3yr = FindElement(By.XPath(path_3yr)).Text.Trim(new Char[] { '+' });
                    FindElement(By.XPath("//li[@class='mod-ui-tab mod-ui-tab__module-header'][contains(text(),'Annual')]")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    path_ytd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Focused Equity Opportunities Fund Class M Shares')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    value_ytd = FindElement(By.XPath(path_ytd)).Text.Trim(new Char[] { '+' });
                    break;
                case 14:
                    path_mtd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Mid Cap Multi-Strategy Fund Class M')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    path_3mo = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Mid Cap Multi-Strategy Fund Class M')]/parent::td" +
                        "/following-sibling::td[5]/child::span";
                    path_1yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Mid Cap Multi-Strategy Fund Class M')]/parent::td" +
                        "/following-sibling::td[3]/child::span";
                    path_3yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Mid Cap Multi-Strategy Fund Class M')]/parent::td" +
                        "/following-sibling::td[2]/child::span";
                    value_mtd = FindElement(By.XPath(path_mtd)).Text.Trim(new Char[] { '+' });
                    value_3mo = FindElement(By.XPath(path_3mo)).Text.Trim(new Char[] { '+' });
                    value_1yr = FindElement(By.XPath(path_1yr)).Text.Trim(new Char[] { '+' });
                    value_3yr = FindElement(By.XPath(path_3yr)).Text.Trim(new Char[] { '+' });
                    FindElement(By.XPath("//li[@class='mod-ui-tab mod-ui-tab__module-header'][contains(text(),'Annual')]")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    path_ytd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Mid Cap Multi-Strategy Fund Class M')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    value_ytd = FindElement(By.XPath(path_ytd)).Text.Trim(new Char[] { '+' });
                    break;
                case 15:
                    path_mtd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Select Managers Small Cap Value Fund Class Y')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    path_3mo = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Select Managers Small Cap Value Fund Class Y')]/parent::td" +
                        "/following-sibling::td[5]/child::span";
                    path_1yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Select Managers Small Cap Value Fund Class Y')]/parent::td" +
                        "/following-sibling::td[3]/child::span";
                    path_3yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Select Managers Small Cap Value Fund Class Y')]/parent::td" +
                        "/following-sibling::td[2]/child::span";
                    value_mtd = FindElement(By.XPath(path_mtd)).Text.Trim(new Char[] { '+' });
                    value_3mo = FindElement(By.XPath(path_3mo)).Text.Trim(new Char[] { '+' });
                    value_1yr = FindElement(By.XPath(path_1yr)).Text.Trim(new Char[] { '+' });
                    value_3yr = FindElement(By.XPath(path_3yr)).Text.Trim(new Char[] { '+' });
                    FindElement(By.XPath("//li[@class='mod-ui-tab mod-ui-tab__module-header'][contains(text(),'Annual')]")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    path_ytd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Select Managers Small Cap Value Fund Class Y')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    value_ytd = FindElement(By.XPath(path_ytd)).Text.Trim(new Char[] { '+' });
                    break;
                case 16:
                    path_mtd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Select Managers Small Cap Growth Fund Class Y')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    path_3mo = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Select Managers Small Cap Growth Fund Class Y')]/parent::td" +
                        "/following-sibling::td[5]/child::span";
                    path_1yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Select Managers Small Cap Growth Fund Class Y')]/parent::td" +
                        "/following-sibling::td[3]/child::span";
                    path_3yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Select Managers Small Cap Growth Fund Class Y')]/parent::td" +
                        "/following-sibling::td[2]/child::span";
                    value_mtd = FindElement(By.XPath(path_mtd)).Text.Trim(new Char[] { '+' });
                    value_3mo = FindElement(By.XPath(path_3mo)).Text.Trim(new Char[] { '+' });
                    value_1yr = FindElement(By.XPath(path_1yr)).Text.Trim(new Char[] { '+' });
                    value_3yr = FindElement(By.XPath(path_3yr)).Text.Trim(new Char[] { '+' });
                    FindElement(By.XPath("//li[@class='mod-ui-tab mod-ui-tab__module-header'][contains(text(),'Annual')]")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    path_ytd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Select Managers Small Cap Growth Fund Class Y')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    value_ytd = FindElement(By.XPath(path_ytd)).Text.Trim(new Char[] { '+' });
                    break;
                case 17:
                    path_mtd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon International Stock Fund Class Y')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    path_3mo = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon International Stock Fund Class Y')]/parent::td" +
                        "/following-sibling::td[5]/child::span";
                    path_1yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon International Stock Fund Class Y')]/parent::td" +
                        "/following-sibling::td[3]/child::span";
                    path_3yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon International Stock Fund Class Y')]/parent::td" +
                        "/following-sibling::td[2]/child::span";
                    value_mtd = FindElement(By.XPath(path_mtd)).Text.Trim(new Char[] { '+' });
                    value_3mo = FindElement(By.XPath(path_3mo)).Text.Trim(new Char[] { '+' });
                    value_1yr = FindElement(By.XPath(path_1yr)).Text.Trim(new Char[] { '+' });
                    value_3yr = FindElement(By.XPath(path_3yr)).Text.Trim(new Char[] { '+' });
                    FindElement(By.XPath("//li[@class='mod-ui-tab mod-ui-tab__module-header'][contains(text(),'Annual')]")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    path_ytd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon International Stock Fund Class Y')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    value_ytd = FindElement(By.XPath(path_ytd)).Text.Trim(new Char[] { '+' });
                    break;
                case 18:
                    path_mtd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon International Equity Fund Class Y')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    path_3mo = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon International Equity Fund Class Y')]/parent::td" +
                        "/following-sibling::td[5]/child::span";
                    path_1yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon International Equity Fund Class Y')]/parent::td" +
                        "/following-sibling::td[3]/child::span";
                    path_3yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon International Equity Fund Class Y')]/parent::td" +
                        "/following-sibling::td[2]/child::span";
                    value_mtd = FindElement(By.XPath(path_mtd)).Text.Trim(new Char[] { '+' });
                    value_3mo = FindElement(By.XPath(path_3mo)).Text.Trim(new Char[] { '+' });
                    value_1yr = FindElement(By.XPath(path_1yr)).Text.Trim(new Char[] { '+' });
                    value_3yr = FindElement(By.XPath(path_3yr)).Text.Trim(new Char[] { '+' });
                    FindElement(By.XPath("//li[@class='mod-ui-tab mod-ui-tab__module-header'][contains(text(),'Annual')]")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    path_ytd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon International Equity Fund Class Y')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    value_ytd = FindElement(By.XPath(path_ytd)).Text.Trim(new Char[] { '+' });
                    break;
                case 19:
                    path_mtd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Emerging Markets Fund Class M')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    path_3mo = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Emerging Markets Fund Class M')]/parent::td" +
                        "/following-sibling::td[5]/child::span";
                    path_1yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Emerging Markets Fund Class M')]/parent::td" +
                        "/following-sibling::td[3]/child::span";
                    path_3yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Emerging Markets Fund Class M')]/parent::td" +
                        "/following-sibling::td[2]/child::span";
                    value_mtd = FindElement(By.XPath(path_mtd)).Text.Trim(new Char[] { '+' });
                    value_3mo = FindElement(By.XPath(path_3mo)).Text.Trim(new Char[] { '+' });
                    value_1yr = FindElement(By.XPath(path_1yr)).Text.Trim(new Char[] { '+' });
                    value_3yr = FindElement(By.XPath(path_3yr)).Text.Trim(new Char[] { '+' });
                    FindElement(By.XPath("//li[@class='mod-ui-tab mod-ui-tab__module-header'][contains(text(),'Annual')]")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    path_ytd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'BNY Mellon Emerging Markets Fund Class M')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    value_ytd = FindElement(By.XPath(path_ytd)).Text.Trim(new Char[] { '+' });
                    break;
                case 20:
                    path_mtd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'TIAA-CREF High-Yield Fund Institutional Class')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    path_3mo = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'TIAA-CREF High-Yield Fund Institutional Class')]/parent::td" +
                        "/following-sibling::td[5]/child::span";
                    path_1yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'TIAA-CREF High-Yield Fund Institutional Class')]/parent::td" +
                        "/following-sibling::td[3]/child::span";
                    path_3yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'TIAA-CREF High-Yield Fund Institutional Class')]/parent::td" +
                        "/following-sibling::td[2]/child::span";
                    value_mtd = FindElement(By.XPath(path_mtd)).Text.Trim(new Char[] { '+' });
                    value_3mo = FindElement(By.XPath(path_3mo)).Text.Trim(new Char[] { '+' });
                    value_1yr = FindElement(By.XPath(path_1yr)).Text.Trim(new Char[] { '+' });
                    value_3yr = FindElement(By.XPath(path_3yr)).Text.Trim(new Char[] { '+' });
                    FindElement(By.XPath("//li[@class='mod-ui-tab mod-ui-tab__module-header'][contains(text(),'Annual')]")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    path_ytd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'TIAA-CREF High-Yield Fund Institutional Class')]/parent::td" +
                        "/following-sibling::td[6]/child::span";
                    value_ytd = FindElement(By.XPath(path_ytd)).Text.Trim(new Char[] { '+' });
                    break;
            }
            string[] bp_vals = { value_mtd, value_3mo, value_ytd, value_1yr, value_3yr };
            return bp_vals;
        }
    }
}
