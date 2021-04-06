﻿using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Forms;

namespace C_sharp_MSEdge_Chromium_Browser_automating
{
    class Browser
    {
        Browser(BrowserName browserName)
        {

        }

        /*fredrikhaglund/ChromeLauncher.cs
        https://gist.github.com/fredrikhaglund/43aea7522f9e844d3e7b
         */
        private const string ChromeAppKey = 
            @"\Software\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe";

        private static string ChromeAppFileName
        {
            get
            {
                return (string)(Registry.GetValue("HKEY_LOCAL_MACHINE" + ChromeAppKey, "", null) ??
                                    Registry.GetValue("HKEY_CURRENT_USER" + ChromeAppKey, "", null));
            }
        }

        public static void OpenLink(string url)
        {
            string chromeAppFileName = ChromeAppFileName;
            if (string.IsNullOrEmpty(chromeAppFileName))
            {
                throw new Exception("Could not find chrome.exe!");
            }
            Process.Start(chromeAppFileName, url);
        }


        public static string[] getUrl(BrowserName browserNameFrom)
        {//https://www.c-sharpcorner.com/forums/how-to-all-the-urls-of-the-open-tabs-of-a-browser
            string[] msg={ "",""};
            string browsername = "chrome";
            switch (browserNameFrom)
            {
                case BrowserName.Chrome:
                    break;
                case BrowserName.MsEdge:
                    browsername = "msedge";
                    break;
                default:
                    break;
            }
            try
            {
                //Process[] procsChrome = Process.GetProcessesByName("chrome");
                Process[] procsChrome = Process.GetProcessesByName(browsername);
                if (procsChrome.Length <= 0)
                {
                    MessageBox.Show("Chrome is not running");
                }
                else
                {
                    string urls = "";
                    foreach (Process proc in procsChrome)
                    {
                        // the chrome process must have a window
                        if (proc.MainWindowHandle == IntPtr.Zero)
                        {
                            continue;
                        }

                        // find the automation element
                        AutomationElement elm = AutomationElement.FromHandle
                            (proc.MainWindowHandle);
                        //AutomationElement elmUrlBar =
                        //    elm.FindFirst(TreeScope.Descendants,
                        //    new PropertyCondition(AutomationElement.NameProperty,
                        //    "Address and search bar"));
                        AutomationElementCollection elmUrlBar =
                            elm.FindAll(TreeScope.Subtree,
                            new PropertyCondition(
                                AutomationElement.ControlTypeProperty,
                                ControlType.Edit));//https://social.msdn.microsoft.com/Forums/en-US/f9cb8d8a-ab6e-4551-8590-bda2c38a2994/retrieve-chrome-url-using-automation-element-in-c-application?forum=csharpgeneral
                        /*要用Edit屬性才抓得到網址列,Text也不行
                         */

                        // if it can be found, get the value from the URL bar
                        if (elmUrlBar != null)
                        {
                            foreach (AutomationElement Elm in elmUrlBar)
                            {
                                string vp = ((ValuePattern)Elm.
                                    GetCurrentPattern(ValuePattern.Pattern)).
                                    Current.Value as string;
                                urls += vp;
                            }
                        }
                    }
                    //textBox1.Text = urls;
                    msg[0] = urls;
                    if (urls.IndexOf("https://") > -1)
                    {
                        openUrlChrome(urls);
                    }
                    else
                        openUrlChrome(@"https://" + urls);
                }
            }
            catch (Exception ex)
            {
                //textBox2.Text = ex.ToString();
                msg[1]= ex.ToString();
            }
            return msg;
        }
        void getUrl_noWork()
        //https://stackoverflow.com/questions/18897070/getting-the-current-tabs-url-from-google-chrome-using-c-sharp
        { // there are always multiple chrome processes, so we have to loop through all of them to find the
          // process with a Window Handle and an automation element of name "Address and search bar"
            Process[] procsChrome = Process.GetProcessesByName("chrome");
            string urls = "";
            foreach (Process chrome in procsChrome)
            {
                // the chrome process must have a window
                if (chrome.MainWindowHandle == IntPtr.Zero)
                {
                    continue;
                }

                // find the automation element
                AutomationElement elm = AutomationElement.FromHandle
                    (chrome.MainWindowHandle);
                AutomationElement elmUrlBar =
                    elm.FindFirst(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.NameProperty,
                    "Address and search bar"));
                /*NameProperty 這個屬性抓不到
                 * AutomationElement.ControlTypeProperty,
                ControlType.Edit));//這個個屬性才抓得到網址列，詳 getUrl()
                */

                // if it can be found, get the value from the URL bar
                if (elmUrlBar != null)
                {
                    AutomationPattern[] patterns = elmUrlBar.GetSupportedPatterns();
                    if (patterns.Length > 0)
                    {
                        ValuePattern val =
                            (ValuePattern)elmUrlBar.GetCurrentPattern(patterns[0]);
                        //Console.WriteLine("Chrome URL found: " + val.Current.Value);
                        urls += val.Current.Value;
                    }
                }
            }
            
        }

        static void openUrlChrome(string url)
        {//https://stackoverflow.com/questions/6305388/how-to-launch-a-google-chrome-tab-with-specific-url-using-c-sharp
            //string url = @"https://stackoverflow.com/questions/6305388/how-to-launch-a-google-chrome-tab-with-specific-url-using-c-sharp/";
            string browserFullname = @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe";
            //之前可能是用到WPF所以不接受路徑中有空格，且又有存取權限的問題。這個Windows Forms應用程式則似乎都又有這樣的問題了
            //string browserFullname = @"C:\""Program Files (x86)""\Google\Chrome\Application\google_translation-ConsoleApp.exe";
            //使用空格的長檔名或路徑需要用引號括住:
            //https://docs.microsoft.com/zh-tw/troubleshoot/windows-server/deployment/filenames-with-spaces-require-quotation-mark
            //browserFullname = @"V:\softwares\PortableApps\PortableApps\GoogleChromePortable\GoogleChromePortable.exe";

            Process.Start(browserFullname, url);
            //Process.Start(url);//這樣是用系統預設瀏覽器開啟
        }

        #region MyTempRegion

        string browserFullname = getBrowserFullname(BrowserName.MsEdge);
        private static string getBrowserFullname(BrowserName browserName)
        {//https://stackoverflow.com/questions/14299382/getting-chrome-and-firefox-version-locally-c-sharp
            object path;string bFullname="";
            switch (browserName)
            {
                case BrowserName.Chrome:
                    path = Registry.GetValue
                        (@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe", "", null);
                    if (path != null)
                        bFullname = FileVersionInfo.GetVersionInfo(path.ToString()).FileVersion;
                    else
                        bFullname = "";
                    break;
                case BrowserName.MsEdge:
                    bFullname = "";
                    break;
                default:
                    bFullname = "";
                    break;
            }
            return bFullname;

            //path = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\firefox.exe", "", null);
            //if (path != null)
            //    Console.WriteLine("Firefox: " + FileVersionInfo.GetVersionInfo(path.ToString()).FileVersion);
        }
        #endregion
    }
}