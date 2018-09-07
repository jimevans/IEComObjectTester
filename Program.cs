using MSHTML;
using SHDocVw;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEComObjectTester
{
    class Program
    {
        static Dictionary<string, InternetExplorer> instances = new Dictionary<string, InternetExplorer>();

        static void Main(string[] args)
        {
            string baseUrl = "http://localhost:2310/common/";
            int maxIterations = 50;
            string instanceId = CreateIEInstance();
            var ie = instances[instanceId];
            IHTMLDocument3 doc = null;
            for (int i = 0; i < maxIterations; i++)
            {
                Console.WriteLine("Iteration {0} of {1}", i + 1, maxIterations);
                ie.Navigate2(baseUrl + "xhtmlTest.html");
                System.Threading.Thread.Sleep(500);
                if (doc == null)
                {
                    doc = ie.Document;
                }

                IHTMLElement element = doc.getElementsByName("windowOne").item(0);
                element.click();
                while (instances.Keys.Count < 2)
                {
                    System.Threading.Thread.Sleep(100);
                }

                string popupId = instances.Keys.Except(new List<string>() { instanceId }).First();

                CloseIEInstance(popupId);
                ie.Navigate2(baseUrl + "macbeth.html");
            }

            CloseIEInstance(instanceId);
            Console.WriteLine("Completed {0} iterations. Press <Enter> to end.", maxIterations);
            Console.ReadLine();
        }

        private static void CloseIEInstance(string instanceId)
        {
            instances[instanceId].Quit();
            System.Threading.Thread.Sleep(500);
            instances[instanceId] = null;
            instances.Remove(instanceId);
        }

        private static string CreateIEInstance()
        {
            string instanceId = Guid.NewGuid().ToString();
            var ie = new InternetExplorer();

            ie.Visible = true;
            ie.OnQuit += OnBrowserQuit;
            ie.NewWindow3 += OnNewWindow;
            instances[instanceId] = ie;
            return instanceId;
        }

        private static void OnNewWindow(ref object ppDisp, ref bool Cancel, uint dwFlags, string bstrUrlContext, string bstrUrl)
        {
            string instanceId = CreateIEInstance();
            var ie = instances[instanceId];
            ppDisp = ie;
        }

        private static void OnBrowserQuit()
        {
        }
    }
}
