using System;
using System.Collections.Generic;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Collections.ObjectModel;

/************************************************************************************************
**
**  Date: March 29th, 2021
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
        //Selenium function to pull data from browser for each item in top Dictionary of links
        //Returns string array of new percentages        
        public String[] Pull_perf_sum_data(int l)
        {
            string path1_mtd, path1_3mo, path1_ytd, path1_1yr, path1_3yr, path2_mtd, path2_3mo, path2_ytd, path2_1yr, path2_3yr;
            string value1_mtd="", value1_3mo="", value1_1yr="", value1_ytd = "", value1_3yr = "", value2_mtd="", value2_3mo = "", value2_1yr = "", value2_ytd = "", value2_3yr = "";
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
                        "/following-sibling::td[6]/child::span[@class='mod-format--pos']";
                    path1_3mo = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'iShares Core MSCI International Developed Markets ETF')]/parent::td" +
                        "/following-sibling::td[5]/child::span[@class='mod-format--pos']";
                    path1_1yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'iShares Core MSCI International Developed Markets ETF')]/parent::td" +
                        "/following-sibling::td[3]/child::span[@class='mod-format--pos']";
                    path1_3yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'iShares Core MSCI International Developed Markets ETF')]/parent::td" +
                        "/following-sibling::td[2]/child::span[@class='mod-format--pos']";
                    path2_mtd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'MSCI World Ex USA IMI NR USD')]/parent::td/following-sibling::td[6]" +
                        "/child::span[@class='mod-format--pos']";
                    path2_3mo = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'MSCI World Ex USA IMI NR USD')]/parent::td/following-sibling::td[5]" +
                        "/child::span[@class='mod-format--pos']";                    
                    path2_1yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'MSCI World Ex USA IMI NR USD')]/parent::td/following-sibling::td[3]" +
                        "/child::span[@class='mod-format--pos']";
                    path2_3yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'MSCI World Ex USA IMI NR USD')]/parent::td/following-sibling::td[2]" +
                        "/child::span[@class='mod-format--pos']";                    
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
                        "/following-sibling::td[6]/child::span[@class='mod-format--pos']";
                    path2_ytd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'MSCI World Ex USA IMI NR USD')]/parent::td/following-sibling::td[6]" +
                        "/child::span[@class='mod-format--pos']";
                    value1_ytd = FindElement(By.XPath(path1_ytd)).Text.Trim(new Char[] { '+' });
                    value2_ytd = FindElement(By.XPath(path2_ytd)).Text.Trim(new Char[] { '+' });
                    break;
                case 4:
                    path1_mtd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'Goldman Sachs ActiveBeta® Emerging Markets Equity ETF')]/parent::td" +
                        "/following-sibling::td[6]/child::span[@class='mod-format--pos']";
                    path1_3mo = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'Goldman Sachs ActiveBeta® Emerging Markets Equity ETF')]/parent::td" +
                        "/following-sibling::td[5]/child::span[@class='mod-format--pos']";
                    path1_1yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'Goldman Sachs ActiveBeta® Emerging Markets Equity ETF')]/parent::td" +
                        "/following-sibling::td[3]/child::span[@class='mod-format--pos']";
                    path1_3yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'Goldman Sachs ActiveBeta® Emerging Markets Equity ETF')]/parent::td" +
                        "/following-sibling::td[2]/child::span[@class='mod-format--pos']";
                    path2_mtd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'Diversified Emerging Mkts')]/parent::td/following-sibling::td[6]" +
                        "/child::span[@class='mod-format--pos']";
                    path2_3mo = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'Diversified Emerging Mkts')]/parent::td/following-sibling::td[5]" +
                        "/child::span[@class='mod-format--pos']";
                    path2_1yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'Diversified Emerging Mkts')]/parent::td/following-sibling::td[3]" +
                        "/child::span[@class='mod-format--pos']";
                    path2_3yr = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'Diversified Emerging Mkts')]/parent::td/following-sibling::td[2]" +
                        "/child::span[@class='mod-format--pos']";                    
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
                        "/following-sibling::td[6]/child::span[@class='mod-format--pos']";
                    value1_ytd = FindElement(By.XPath(path1_ytd)).Text.Trim(new Char[] { '+' });
                    path2_ytd = "//span[@class='mod-ui-table__cell--colored__wrapper'][contains(text(),'Diversified Emerging Mkts')]/parent::td/following-sibling::td[6]" +
                        "/child::span[@class='mod-format--pos']";
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
                    FindElement(By.XPath("//button[@id='onetrust-accept-btn-handler']")).Click();
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
    }
}
