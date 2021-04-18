using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using Microsoft.Win32;

namespace C_sharp_MSEdge_Chromium_Browser_automating
{
    public class BrowserChrome
    {
        #region fredrikhaglund/ChromeLauncher.cs
        /*fredrikhaglund/ChromeLauncher.cs
        https://gist.github.com/fredrikhaglund/43aea7522f9e844d3e7b
         */
        private const string ChromeAppKey =
            @"\Software\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe";

        public static string ChromeAppFileName
        {
            get
            {
                return (string)(Registry.GetValue("HKEY_LOCAL_MACHINE" +
                    ChromeAppKey, "", null) ??
                    Registry.GetValue("HKEY_CURRENT_USER" + ChromeAppKey,
                    "", null));
            }
        }

        public static void OpenLinkChrome(string url)
        {
            string chromeAppFileName = ChromeAppFileName;
            if (string.IsNullOrEmpty(chromeAppFileName))
            {
                throw new Exception("Could not find chrome.exe!");
            }
            Process.Start(chromeAppFileName, Browser.urlRegx(url));
        }
        #endregion

        /*
        string urlRegx(string url)
        {//網址規範化-將特殊字元置換,並清除不必要之字元
            string[] replWds = { "\"", "%22"," ","%20" };//, "http//", "" };
            for (int i = 0; i < replWds.Length; i++)
            {
                url = url.Replace(replWds[i], replWds[++i]);
            }
            #region HTTP not HTTPs
            //List<string> webSitesHTTP = new List<string> { "dict.revised.moe.edu.tw" };
            //foreach (string websitehttp in webSitesHTTP)
            //{
            //    if (url.IndexOf(websitehttp) > -1)
            //    {
            //        url = url.Replace("https://", "http://");
            //    }

            //}
            #endregion
            return url;//url.Replace("\"", "%22");
        }
        */
    }
}
