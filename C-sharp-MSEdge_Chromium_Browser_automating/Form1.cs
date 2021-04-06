using System;
using System.Diagnostics;
using System.Windows.Automation;
using System.Windows.Forms;

//https://social.msdn.microsoft.com/Forums/en-US/7bdafd2a-be91-4f4f-a33d-6bea2f889e09/c-sample-for-automating-ms-edge-chromium-browser-using-edge-web-driver?forum=iewebdevelopment
//using OpenQA.Selenium;
//using OpenQA.Selenium.Chrome;
//using OpenQA.Selenium.Edge;
//using OpenQA.Selenium.Remote;
//using OpenQA.Selenium.Support.UI;

namespace C_sharp_MSEdge_Chromium_Browser_automating
{
    public partial class Form1 : Form
    {
     


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //https://docs.microsoft.com/en-us/microsoft-edge/webview2/
            //★★★ https://docs.microsoft.com/en-us/microsoft-edge/webdriver-chromium/?tabs=c-sharp
            //https://stackoverflow.com/questions/39626509/how-to-launch-ms-edge-from-c-sharp-winforms#_=_
            //https://www.google.com/search?q=C%23+msedge&oq=C%23+msedge&aqs=edge..69i57j69i58j0i8i30l2.9300j0j4&sourceid=chrome&ie=UTF-8
            //selenium 
            //https://medium.com/@geekanonymous79/microsoft-edge-automation-using-selenium-in-python-af3cecfed5ed
            
            string[] msg=new Browser(BrowserName.MsEdge).getUrl();
            textBox1.Text = msg[0];
            textBox2.Text = msg[1];
        }


        public void getTabTitle()//Checkurl()
        {//https://www.c-sharpcorner.com/forums/how-to-all-the-urls-of-the-open-tabs-of-a-browser
            try
            {
                Process[] procsChrome = Process.GetProcessesByName("chrome");
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
                        AutomationElement root = AutomationElement.FromHandle(proc.MainWindowHandle);
                        Condition condition = new PropertyCondition
                            (AutomationElement.ControlTypeProperty,
                            ControlType.TabItem);//以此TabItem屬性找只能找到分頁的title值，要用Edit才能找到網址列之值（詳getUrl()）
                        var tabs = root.FindAll(TreeScope.Descendants, condition);
                        foreach (AutomationElement tabitem in tabs)
                        {
                            string urlname = tabitem.Current.Name;
                            urls = urls + "\r\n" + urlname;
                        }
                    }
                    textBox1.Text = urls;
                }
            }
            catch (Exception ex)
            {
                textBox2.Text = ex.ToString();
            }
        }


        
    }

    public enum BrowserName
    {
        Chrome, MsEdge
    }
}


//https://nugetmusthaves.com/Tag/selenium
//https://social.msdn.microsoft.com/Forums/en-US/7bdafd2a-be91-4f4f-a33d-6bea2f889e09/c-sample-for-automating-ms-edge-chromium-browser-using-edge-web-driver?forum=iewebdevelopment
//    var anaheimService = ChromeDriverService.CreateDefaultService(@"D:\", "

//    msedgedriver.exe "); 
//       // user need to pass the driver path here....
//       var anaheimOptions = new ChromeOptions
//       {
//           // user need to pass the location of new edge app here....
//           BinaryLocation = @"
// C: \Program Files(x86)\Microsoft\Edge Dev\Application\msedge.exe "
//       };

//    IWebDriver driver = new ChromeDriver(anaheimService, anaheimOptions);
//    driver.Navigate().GoToUrl("https: //example.com/");
//    Console.WriteLine(driver.Title.ToString());
//    driver.Close();