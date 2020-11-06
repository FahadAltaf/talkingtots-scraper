using CsvHelper;
using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace talkingtots_scraper
{
    public class DataModel
    {
        public string FullName { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }
        public string AssignedFranchise { get; set; }
        public string Active { get; set; }
        public string DateSubscribed { get; set; }
        public string TelePhone { get; set; }
        public string RecieveEmail { get; set; }
        public string RecieveSms { get; set; }
        public string Category { get; set; }
    }
    class Program
    {
        static List<DataModel> entries = new List<DataModel>();
        static void Main(string[] args)
        {
            using (var driver = new ChromeDriver())
            {
                driver.Navigate().GoToUrl("https://www.talkingtots.info/admin");

                Console.WriteLine("Login and press enter");
                Console.ReadKey();

                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(driver.PageSource);

                var table = doc.DocumentNode.SelectSingleNode("//*[@id=\"form1\"]/table/tbody/tr/td[2]/table/tbody/tr[4]/td/table/tbody");
                var rows = table.ChildNodes.Where(x => x.Name == "tr").Skip(1).ToList();
                foreach (var row in rows)
                {
                    try
                    {
                        var entry = new DataModel();
                        var index = rows.IndexOf(row);
                        var sub = new HtmlDocument();
                        sub.LoadHtml(row.InnerHtml);

                        entry.FullName = sub.DocumentNode.SelectSingleNode("/td[2]").InnerText.Replace("\r\n  ", "").Trim();
                        entry.Email = sub.DocumentNode.SelectSingleNode("/td[3]").InnerText.Replace("\r\n  ", "").Trim();
                        entry.AssignedFranchise = sub.DocumentNode.SelectSingleNode("/td[4]").InnerText.Replace("\r\n  ", "").Trim();
                        entry.Active = sub.DocumentNode.SelectSingleNode("/td[5]").InnerText.Replace("\r\n  ", "").Trim();

                        IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                        js.ExecuteScript("document.getElementById(\"ContentPlaceHolder1_rpTalkTalk_btnManage_" + index + "\").dispatchEvent(new MouseEvent(\"click\", {ctrlKey: true}));");

                        driver.SwitchTo().Window(driver.WindowHandles.Last());

                        //Wait for page to load
                        bool loaded = false;
                        while (!loaded)
                        {
                            try
                            {
                                driver.FindElement(By.Id("ContentPlaceHolder1_lbFirstName")).Click();
                                loaded = true;
                            }
                            catch (Exception)
                            {

                            }

                        }

                        sub.LoadHtml(driver.PageSource);
                        try
                        {
                            entry.FirstName = sub.DocumentNode.SelectSingleNode("//input[@id='ContentPlaceHolder1_lbFirstName']").Attributes.FirstOrDefault(x => x.Name == "value").Value;
                        }
                        catch { }

                        try
                        {
                            entry.LastName = sub.DocumentNode.SelectSingleNode("//input[@id='ContentPlaceHolder1_lbSurname']").Attributes.FirstOrDefault(x => x.Name == "value").Value;
                        }
                        catch { }

                        try
                        {
                            entry.DateSubscribed = sub.DocumentNode.SelectSingleNode("//input[@id='ContentPlaceHolder1_lbDateSubscribed']").Attributes.FirstOrDefault(x => x.Name == "value").Value;
                        }
                        catch { }

                        try
                        {
                            entry.TelePhone = sub.DocumentNode.SelectSingleNode("//input[@id='ContentPlaceHolder1_lbTelNo']").Attributes.FirstOrDefault(x => x.Name == "value").Value;
                        }
                        catch { }

                        try
                        {
                            SelectElement ele = new SelectElement(driver.FindElement(By.Id("ContentPlaceHolder1_lbReceiveEmail")));
                            entry.RecieveEmail = ele.SelectedOption.Text;
                        }
                        catch { }

                        try
                        {
                            SelectElement ele1 = new SelectElement(driver.FindElement(By.Id("ContentPlaceHolder1_lbReceiveSMS")));
                            entry.RecieveSms = ele1.SelectedOption.Text;
                        }
                        catch { }

                        Console.WriteLine(entry.FullName);
                        entries.Add(entry);
                        using (var writer = new StreamWriter("result.csv"))
                        using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                        {
                            csv.WriteRecords(entries);
                        }
                        driver.Close();
                        driver.SwitchTo().Window(driver.WindowHandles.First());
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        Console.ReadKey();
                    }
                }

            }

            

            Console.WriteLine("Operation Completed");
        }
    }
}
