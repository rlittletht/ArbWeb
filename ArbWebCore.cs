using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using AxSHDocVw;
using mshtml;
using System.Text.RegularExpressions;

namespace ArbWeb
{
    public partial class ArbWebCore : Form
    {
        object Zero = 0;
        object EmptyString = "";
        private StatusBox.StatusRpt m_srpt;

        public AxSHDocVw.AxWebBrowser AxWeb
        {
            get { return m_axWebBrowser1; }
        }

        public IHTMLDocument  Document      {  get { return (IHTMLDocument)m_axWebBrowser1.Document; } }
        public IHTMLDocument2 Document2     {  get { return (IHTMLDocument2)m_axWebBrowser1.Document; } }
        public IHTMLDocument3 Document3     {  get { return (IHTMLDocument3)m_axWebBrowser1.Document; } }
        public IHTMLDocument4 Document4     {  get { return (IHTMLDocument4)m_axWebBrowser1.Document; } }
        public IHTMLDocument5 Document5     {  get { return (IHTMLDocument5)m_axWebBrowser1.Document; } }

        public ArbWebCore(StatusBox.StatusRpt srpt)
        {
            m_plNewWindow3 = new List<DWebBrowserEvents2_NewWindow3EventHandler>();
            m_plBeforeNav2= new List<DWebBrowserEvents2_BeforeNavigate2EventHandler>();

            m_srpt = srpt;

            InitializeComponent();
        }

        bool m_fNavDone;

        List<DWebBrowserEvents2_NewWindow3EventHandler> m_plNewWindow3;
        List<DWebBrowserEvents2_BeforeNavigate2EventHandler> m_plBeforeNav2;

        public void PushNewWindow3Delegate(DWebBrowserEvents2_NewWindow3EventHandler oDelegate)
        {
            m_plNewWindow3.Add(oDelegate);
            m_axWebBrowser1.NewWindow3 += oDelegate;
        }

        public void PopNewWindow3Delegate()
        {
            if (m_plNewWindow3.Count <= 0)
                return;

            m_axWebBrowser1.NewWindow3 -= m_plNewWindow3[m_plNewWindow3.Count - 1];
            m_plNewWindow3.RemoveAt(m_plNewWindow3.Count - 1);
        }

        byte[] m_rgbPostData;
        bool m_fNavIntercept;
        string m_sNavIntercept;
        
        public void SaveToFileDelegate(object sender, DWebBrowserEvents2_BeforeNavigate2Event e)
        {
        e.cancel = true;
//#if no
            if (m_fDontIntercept)
                return;
                
        byte[] rgb = (byte[])e.postData;
        
        m_rgbPostData = new byte[rgb.Length];
        
        rgb.CopyTo(m_rgbPostData, 0);
        m_fNavIntercept = true;
        m_sNavIntercept = (string)e.uRL;
        m_fNavDone = true;
//#endif // 0        
        
        //            System.Net.WebClient wc = new System.Net.WebClient();
            
//            wc.DownloadFile((string)e.uRL, m_sSaveToFileTarget);
//            e.cancel = true;
#if no                    
            System.Net.WebRequest wrq;
            System.Net.WebResponse wrs;
            
            wrq = System.Net.HttpWebRequest.Create((string)e.uRL);
            wrq.Method = "GET";
            wrs = wrq.GetResponse();
            
            System.IO.StreamReader sr = new System.IO.StreamReader(wrs.GetResponseStream());
            
            string sResult = sr.ReadToEnd();
            sr.Close();
#endif //
#if no
                        System.IO.StreamReader sr;
                        string sResult ;
                        
            System.Net.HttpWebRequest wrqHttp = (System.Net.HttpWebRequest)System.Net.HttpWebRequest.Create((string)e.uRL);
            wrqHttp.ContentType = "application/x-www-form-urlencoded";
            
            byte[] rgb = (byte[])e.postData;
            wrqHttp.ContentLength = rgb.Length;
            wrqHttp.Method = "POST";
            System.IO.Stream stm = wrqHttp.GetRequestStream();
            
            stm.Write(rgb, 0, rgb.Length);
            stm.Close();
            System.Net.CookieContainer cookies = new System.Net.CookieContainer();
            
            string[] rgsCookies = ((IHTMLDocument2)m_axWebBrowser1.Document).cookie.Split(';');
            
            foreach (string sCookie in rgsCookies)
                cookies.SetCookies(new Uri("https://www.thearbiter.net"), sCookie.Trim());

            wrqHttp.CookieContainer = cookies;
                            
            System.Net.HttpWebResponse rspHttp = (System.Net.HttpWebResponse)wrqHttp.GetResponse();

            if (wrqHttp.CookieContainer != null)
                rspHttp.Cookies = wrqHttp.CookieContainer.GetCookies(rspHttp.ResponseUri);
            
            stm = rspHttp.GetResponseStream();
            
            System.IO.StreamWriter sw = new System.IO.StreamWriter(m_sSaveToFileTarget, false);
            sr = new System.IO.StreamReader(stm);
            char[] rgbCopy = new char[4096];
            int cb;
            while (!sr.EndOfStream)
                {
                cb = sr.ReadBlock(rgbCopy, 0, 4096);
                sw.Write(rgbCopy, 0, cb);
                }
            
            
            sr.Close();
            sw.Close();
            stm.Close();
            rspHttp.Close();
            m_fNavDone = true;
#endif            
        }   
  
        string m_sSaveToFileTarget;
        
        public void dfeh(object sender, DWebBrowserEvents2_FileDownloadEvent fdeh)
        {
            string sTarget = fdeh.activeDocument.ToString();
            }
            
            
        public void PushSaveToFile(string sTargetFile)
        {
//            m_axWebBrowser1.FileDownload += new DWebBrowserEvents2_FileDownloadEventHandler(dfeh);
            m_sSaveToFileTarget = sTargetFile;
            m_plBeforeNav2.Add(new DWebBrowserEvents2_BeforeNavigate2EventHandler(SaveToFileDelegate));
            m_axWebBrowser1.BeforeNavigate2 += m_plBeforeNav2[m_plBeforeNav2.Count - 1];
        }
        
        public void PopSaveToFile()
        {
            m_sSaveToFileTarget = null;
            if (m_plBeforeNav2.Count <= 0)
                return;

            m_axWebBrowser1.BeforeNavigate2 -= m_plBeforeNav2[m_plBeforeNav2.Count - 1];
            m_plBeforeNav2.RemoveAt(m_plBeforeNav2.Count - 1);
        }
            
        public void ResetNav()
        {
            m_fNavDone = false;
        }

        private void TriggerDocumentDone(object sender, AxSHDocVw.DWebBrowserEvents2_DocumentCompleteEvent e)
        {
//            m_srpt.AddMessage("DocumentComplete");
            m_fNavDone = true;
        }

        void WaitForBrowserReady()
        {
            long s = 0;

            while (s < 200 && m_axWebBrowser1.Busy)
            {
                Application.DoEvents();
                System.Threading.Thread.Sleep(1000);
                s++;
            }

            if (s > 20000)
            {
                m_axWebBrowser1.Stop();
                m_axWebBrowser1.Visible = true;
                throw new Exception("browser can't get unbusy");
                
            }
            return;
        }
                
        public bool FNavToPage(string sUrl)
        {
            ReportNavState("Entering FNavToPage: ");

            m_axWebBrowser1.Stop();
            WaitForBrowserReady();
            m_fNavDone = false;
            m_axWebBrowser1.Navigate(sUrl, ref Zero, ref EmptyString, ref EmptyString, ref EmptyString);
            m_axWebBrowser1.Visible = true;

            return FWaitForNavFinish();
        }

        public void RefreshPage()
        {
            m_fNavDone = false;
            m_axWebBrowser1.Refresh();
            FWaitForNavFinish();
        }
        bool m_fDontIntercept;

        public void ReportNavState(string sTag)
        {
//            m_srpt.AddMessage(String.Format("{0}: Busy: {1}, State: {2}, m_fNavDone: {3}", sTag, m_axWebBrowser1.Busy, m_fNavDone, m_axWebBrowser1.ReadyState));
        }

        public bool FWaitForNavFinish()
        {
            long s = 0;

            // ok, always yield and allow it to run first (so things can get pumping)
            Application.DoEvents();
            System.Threading.Thread.Sleep(50);
            Application.DoEvents();

            ReportNavState("Entering WaitForNavFinish: ");
            WaitForBrowserReady();
            while (s < 200 && !m_fNavDone)
            {
                Application.DoEvents();
                if (m_axWebBrowser1.ReadyState == SHDocVw.tagREADYSTATE.READYSTATE_COMPLETE) // m_fNavDone || 
                    break;
                System.Threading.Thread.Sleep(1000);
                s++;
            }

            ReportNavState("After NavDone Loop: ");

            if (m_fNavIntercept)
                {
                m_axWebBrowser1.Stop();
                WaitForBrowserReady();
                
                m_fNavDone = false;
                m_fNavIntercept = false;
                m_fDontIntercept = true;
                object o = m_rgbPostData;
                object sHdr = "Content-Type: application/x-www-form-urlencoded" + Environment.NewLine;
                object oHdr = sHdr;
                int flags = 0x1;
                object oFlags = flags;
                string sTgt = "_SELF";
                object oTgt = sTgt;
                
                m_axWebBrowser1.Navigate(m_sNavIntercept, ref Zero, ref oTgt, ref o, ref oHdr);
                bool f = FWaitForNavFinish();
                m_fDontIntercept = false;
                
                return f;
                }
                
            if (s > 20000)
            {
                m_axWebBrowser1.Stop();
                m_axWebBrowser1.Visible = true;
                return false;
            }
            return true;
        }

        private void BeforeNav2(object sender, DWebBrowserEvents2_BeforeNavigate2Event e)
        {
//            m_srpt.AddMessage("BeforeNav2");
        }

        private void NavComplete2(object sender, DWebBrowserEvents2_NavigateComplete2Event e)
        {
//            m_srpt.AddMessage("NavComplete2");
        }

        private void DownloadBegin(object sender, EventArgs e)
        {
//            m_srpt.AddMessage("DownloadBegin");
        }

        private void DownloadComplete(object sender, EventArgs e)
        {
//            m_srpt.AddMessage("DownloadComplete");
        }
    }   
}