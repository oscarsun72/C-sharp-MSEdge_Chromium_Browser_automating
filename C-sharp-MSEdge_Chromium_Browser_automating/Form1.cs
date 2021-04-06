using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

        }

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