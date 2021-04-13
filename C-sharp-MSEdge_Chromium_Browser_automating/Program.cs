using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace C_sharp_MSEdge_Chromium_Browser_automating
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            BrowserName bn = BrowserName.MsEdge;
            Process[] procsBrowser = Process.GetProcessesByName("iexplore");
            if (procsBrowser.Length>0)
            {
                bn = BrowserName.iExplore;
            }
            new Browser(bn).getUrlGo();//直接執行，不啟始表單
            //Application.Run(new Form1());//不啟始表單
        }
    }
}
