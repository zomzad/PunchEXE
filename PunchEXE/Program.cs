using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Microsoft.Win32;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;

namespace PunchEXE
{
    class Program
    {
        private static bool _ieRead;
        private static readonly DateTime NowDt = DateTime.Now;
        private static readonly string Date = NowDt.ToString("yyyyMMdd");
        private static readonly string TimeTag = NowDt.Hour > 12 ? "下午" : "上午";
        private static readonly string Week = NowDt.DayOfWeek.ToString("d");

        static void Main(string[] args)
        {
            var time = int.Parse($"{DateTime.Now.Hour.ToString().PadLeft(2,'0')}{DateTime.Now.Minute.ToString().PadLeft(2,'0')}");
            var path = $@"D:\PunchLog\{Date}\{TimeTag}.txt";

            if ((time >= 0800 && time <= 0830 || time >= 1735 && time <= 1750) && File.Exists(path) == false && new List<string>{"0","6"}.Contains(Week) == false)
            {
                log(DateTime.Now + " : 執行打卡", path);

                var driver = new InternetExplorerDriver();
                driver.Navigate().GoToUrl("http://jshradbs01/AttManual/default.aspx");
                var acc = driver.FindElementById("tAccount");
                var pass = driver.FindElementById("tPASSWORD");
                var okBtn = driver.FindElementById("btnSign");

                Thread.Sleep(1000);
                acc.SendKeys("Ryan9891");
                pass.SendKeys("Mu6gy100rd+_");
                okBtn.Click();
                driver.Quit();

                //寫法2
                //SHDocVw.InternetExplorer ie = new SHDocVw.InternetExplorer();
                //ie.DocumentComplete += ie_DocumentComplete;
                //ie.Navigate("http://jshradbs01/AttManual/default.aspx");
                //ie.Visible = true;
                //System.Threading.Thread.Sleep(1000);
                //mshtml.HTMLDocument doc = ie.Document;

                //doc.getElementById("tAccount").innerText = "Ryan9891";
                //doc.getElementById("tPASSWORD").innerText = "Mu6gy100rd++";
                //doc.getElementById("btnSign").click();

                log(DateTime.Now + " : 打卡完成", path);
                Console.Write("打卡完成");
            }
        }

        private static void log(string message, string path)
        {
            FileInfo finfo = new FileInfo(path);
            if (finfo.Directory.Exists == false)
            {
                finfo.Directory.Create();
            }

            FileStream fs = File.Open(path, FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
            StreamWriter sw = new StreamWriter(fs);
            sw.Write(message + Environment.NewLine);
            sw.Dispose();
            fs.Dispose();
        }

        private static void ie_DocumentComplete(object pDisp, ref object URL)
        {
            _ieRead = true;
        }
    }
}
