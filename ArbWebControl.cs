using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using AxSHDocVw;
using mshtml;
using SHDocVw;
using StatusBox;
using DWebBrowserEvents2_BeforeNavigate2EventHandler = AxSHDocVw.DWebBrowserEvents2_BeforeNavigate2EventHandler;
using DWebBrowserEvents2_NewWindow3EventHandler = AxSHDocVw.DWebBrowserEvents2_NewWindow3EventHandler;

namespace ArbWeb
{
    public partial class ArbWebControl : Form
    {
        object Zero = 0;
        object EmptyString = "";
        private StatusRpt m_srpt;

        public AxWebBrowser AxWeb
        {
            get { return m_wbc; }
        }

        public IHTMLDocument  Document      {  get { return (IHTMLDocument)m_wbc.Document; } }
        public IHTMLDocument2 Document2     {  get { return (IHTMLDocument2)m_wbc.Document; } }
        public IHTMLDocument3 Document3     {  get { return (IHTMLDocument3)m_wbc.Document; } }
        public IHTMLDocument4 Document4     {  get { return (IHTMLDocument4)m_wbc.Document; } }
        public IHTMLDocument5 Document5     {  get { return (IHTMLDocument5)m_wbc.Document; } }

        public ArbWebControl(StatusRpt srpt)
        {
#if notused
            m_plNewWindow3 = new List<DWebBrowserEvents2_NewWindow3EventHandler>();
            m_plBeforeNav2= new List<DWebBrowserEvents2_BeforeNavigate2EventHandler>();
#endif
            m_srpt = srpt;

            InitializeComponent();
        }

        bool m_fNavDone;

#if notused
        List<DWebBrowserEvents2_NewWindow3EventHandler> m_plNewWindow3;
        List<DWebBrowserEvents2_BeforeNavigate2EventHandler> m_plBeforeNav2;

        public void PushNewWindow3Delegate(DWebBrowserEvents2_NewWindow3EventHandler oDelegate)
        {
            m_plNewWindow3.Add(oDelegate);
            m_wbc.NewWindow3 += oDelegate;
        }

        public void PopNewWindow3Delegate()
        {
            if (m_plNewWindow3.Count <= 0)
                return;

            m_wbc.NewWindow3 -= m_plNewWindow3[m_plNewWindow3.Count - 1];
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
            m_wbc.BeforeNavigate2 += m_plBeforeNav2[m_plBeforeNav2.Count - 1];
        }
        
        public void PopSaveToFile()
        {
            m_sSaveToFileTarget = null;
            if (m_plBeforeNav2.Count <= 0)
                return;

            m_wbc.BeforeNavigate2 -= m_plBeforeNav2[m_plBeforeNav2.Count - 1];
            m_plBeforeNav2.RemoveAt(m_plBeforeNav2.Count - 1);
        }
#endif // notused
            
        public void ResetNav()
        {
            m_fNavDone = false;
        }

        private void TriggerDocumentDone(object sender, DWebBrowserEvents2_DocumentCompleteEvent e)
        {
//            m_srpt.AddMessage("DocumentComplete");
            m_fNavDone = true;
        }

        void WaitForBrowserReady()
        {
            long s = 0;

            while (s < 200 && m_wbc.Busy)
            {
                Application.DoEvents();
                Thread.Sleep(50);
                s++;
            }

            if (s > 20000)
            {
                m_wbc.Stop();
                m_wbc.Visible = true;
                throw new Exception("browser can't get unbusy");
                
            }
            return;
        }
                
        public bool FNavToPage(string sUrl)
        {
            ReportNavState("Entering FNavToPage: ");

            m_wbc.Stop();
            WaitForBrowserReady();
            m_fNavDone = false;
            m_wbc.Navigate(sUrl, ref Zero, ref EmptyString, ref EmptyString, ref EmptyString);
            m_wbc.Visible = true;

            return FWaitForNavFinish();
        }

        public void RefreshPage()
        {
            m_fNavDone = false;
            m_wbc.Refresh();
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
            Thread.Sleep(50);
            Application.DoEvents();

            ReportNavState("Entering WaitForNavFinish: ");
            WaitForBrowserReady();
            while (s < 200 && !m_fNavDone)
            {
                Application.DoEvents();
                if (m_wbc.ReadyState == tagREADYSTATE.READYSTATE_COMPLETE) // m_fNavDone || 
                    break;
                Thread.Sleep(50);
                s++;
            }

            ReportNavState("After NavDone Loop: ");

#if notused
            if (m_fNavIntercept)
                {
                m_wbc.Stop();
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
                
                m_wbc.Navigate(m_sNavIntercept, ref Zero, ref oTgt, ref o, ref oHdr);
                bool f = FWaitForNavFinish();
                m_fDontIntercept = false;
                
                return f;
                }
#endif                 
            if (s > 20000)
            {
                m_wbc.Stop();
                m_wbc.Visible = true;
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
        static public bool FSetCheckboxControlVal(IHTMLDocument2 oDoc2, bool fChecked, string sName)
        {
            IHTMLElementCollection hec;

            hec = (IHTMLElementCollection)oDoc2.all.tags("input");
            
            foreach (IHTMLInputElement ihie in hec)
                {
                if (String.Compare(ihie.name, sName, true) == 0)
                    {
                    if (ihie.@checked == fChecked)
                        return false;

                    ihie.@checked = fChecked;
                    return true;
                    }
                }
            return false;
        }

        public static bool FSetTextareaControlText(IHTMLDocument2 oDoc2, string sName, string sValue, bool fCheck)
        {
            IHTMLElementCollection hec;

            hec = (IHTMLElementCollection)oDoc2.all.tags("textarea");
            string sT = null;
            bool fNeedSave = false;
            foreach (IHTMLTextAreaElement ihie in hec)
                {
                if (String.Compare(ihie.name, sName, true) == 0)
                    {
                    if (fCheck)
                        {
                        sT = ihie.value;
                        if (sT == null)
                            sT = "";
                        if (String.Compare(sValue, sT) != 0)
                            fNeedSave = true;
                        }
                    else
                        {
                        fNeedSave = true;
                        }
                    ihie.value = sValue;
                    }
                }
            return fNeedSave;
        }
        public static bool FSetInputControlText(IHTMLDocument2 oDoc2, string sName, string sValue, bool fCheck)
        {
            IHTMLElementCollection hec;

            hec = (IHTMLElementCollection)oDoc2.all.tags("input");
            string sT = null;
            bool fNeedSave = false;
            foreach (IHTMLInputElement ihie in hec)
                {
                if (String.Compare(ihie.name, sName, true) == 0)
                    {
                    if (fCheck)
                        {
                        sT = ihie.value;
                        if (sT == null)
                            sT = "";
                        if (String.Compare(sValue, sT) != 0)
                            fNeedSave = true;
                        }
                    else
                        {
                        fNeedSave = true;
                        }
                    ihie.value = sValue;
                    }
                }
            return fNeedSave;
        }
        /* M P  G E T  S E L E C T  V A L U E S */
		/*----------------------------------------------------------------------------
			%%Function: MpGetSelectValues
			%%Qualified: ArbWeb.AwMainForm.MpGetSelectValues
			%%Contact: rlittle

            for a given <select name=$sName><option value=$sValue>$sText</option>...
         
            Find the given sName select object. Then add a mapping of
            $sText -> $sValue to a dictionary and return it.
		----------------------------------------------------------------------------*/
        public static Dictionary<string, string> MpGetSelectValues(StatusRpt srpt, IHTMLDocument2 oDoc2, string sName)
        {
            IHTMLElementCollection hec;
			
            Dictionary<string, string> mp = new Dictionary<string, string>();

            hec = (IHTMLElementCollection)oDoc2.all.tags("select");
            foreach (HTMLSelectElementClass ihie in hec)
                {
                if (String.Compare(ihie.name, sName, true) == 0)
                    {
                    foreach (IHTMLOptionElement ihoe in (IHTMLElementCollection)ihie.tags("option"))
                        {
                        if (mp.ContainsKey(ihoe.text))
                            srpt.AddMessage(String.Format("How strange!  '{0}' shows up more than once as a position", ihoe.text), StatusRpt.MSGT.Warning);
                        else
                            mp.Add(ihoe.text, ihoe.value);
                        }
                    }
                }
            return mp;
        }

        // if fValueIsValue == false, then sValue is the "text" of the option control
        public static bool FSelectMultiSelectOption(IHTMLDocument2 oDoc2, string sName, string sValue, bool fValueIsValue)
        {
            IHTMLElementCollection hec;

            hec = (IHTMLElementCollection)oDoc2.all.tags("select");

            foreach (HTMLSelectElementClass ihie in hec)
                {
                if (String.Compare(ihie.name, sName, true) == 0)
                    {
                    foreach (IHTMLOptionElement ihoe in (IHTMLElementCollection)ihie.tags("option"))
                        {
                        if ((fValueIsValue && ihoe.value == sValue) ||
                            (!fValueIsValue && ihoe.text == sValue))
                            {
                            ihoe.selected = true;
                            return true;
                            }
                        }
                    }
                }
            return false;
        }
        public static bool FResetMultiSelectOptions(IHTMLDocument2 oDoc2, string sName)
        {
            IHTMLElementCollection hec;

            hec = (IHTMLElementCollection)oDoc2.all.tags("select");

            foreach (HTMLSelectElementClass ihie in hec)
                {
                if (String.Compare(ihie.name, sName, true) == 0)
                    {
                    foreach (IHTMLOptionElement ihoe in (IHTMLElementCollection)ihie.tags("option"))
                        {
                        ihoe.selected = false;
                        }
                    }
                }
            return true;
        }
        public static string SGetFilterID(IHTMLDocument2 oDoc2, string sName, string sValue)
        {
            IHTMLElementCollection hec;

            hec = (IHTMLElementCollection)oDoc2.all.tags("select");
            foreach (HTMLSelectElementClass ihie in hec)
                {
                if (String.Compare(ihie.name, sName, true) == 0)
                    {
                    foreach (IHTMLOptionElement ihoe in (IHTMLElementCollection)ihie.tags("option"))
                        {
                        if (ihoe.text == sValue)
                            {
                            return ihoe.value;
                            }
                        }
                    }
                }
            return null;
        }
        static public bool FSetSelectControlText(IHTMLDocument2 oDoc2, string sName, string sValue, bool fCheck)
        {
            IHTMLElementCollection hec;

            hec = (IHTMLElementCollection)oDoc2.all.tags("select");
            bool fNeedSave = false;
            foreach (IHTMLSelectElement ihie in hec)
                {
                if (String.Compare(ihie.name, sName, true) == 0)
                    {
                    foreach(IHTMLOptionElement ihoe in (IHTMLElementCollection)ihie.tags("option"))
                        {
                        if (ihoe.text == sValue)
                            {
                            // value is already set...
                            if (ihie.value == ihoe.value)
                                return false;
                            ihoe.selected = true;
                            object dummy = null;
                            IHTMLDocument4 oDoc4 = (IHTMLDocument4)oDoc2;
                            object eventObj = oDoc4.CreateEventObject(ref dummy);
                            HTMLSelectElementClass hsec = ihie as HTMLSelectElementClass;
                            hsec.FireEvent("onchange", ref eventObj);
                            return true;
                            }
                        }
                    }
                }
            return fNeedSave;
        }
#if no
		private bool FClickControlName(IHTMLDocument2 oDoc2, string sTag, string sName)
		{
			// find an sTag with name sName
			IHTMLElementCollection hec;
			hec = (IHTMLElementCollection)oDoc2.all.tags(sTag);

			if (sTag.ToUpper() == "a")
				{
				foreach (IHTMLAnchorElement 
				}
		}

#endif // no
        public static bool FClickControl(ArbWebControl awc, IHTMLDocument2 oDoc2, string sId)
        {
//			m_srpt.AddMessage("Before clickcontrol: "+sId);
            ((IHTMLElement)(oDoc2.all.item(sId, 0))).click();
//			m_srpt.AddMessage("After clickcontrol");
            return awc.FWaitForNavFinish();
        }
        static public bool FClickControlNoWait(IHTMLDocument2 oDoc2, string sId)
        {
//			m_srpt.AddMessage("Before clickcontrol: "+sId);
            ((IHTMLElement)(oDoc2.all.item(sId, 0))).click();
//			m_srpt.AddMessage("After clickcontrol");
            return true;
        }
        public static bool FCheckForControl(IHTMLDocument2 oDoc2, string sId)
        {
            if (oDoc2.all.item(sId, 0) != null)
                return true;
                
            return false;
        }
        public static string SGetControlValue(IHTMLDocument2 oDoc2, string sId)
        {
            if (FCheckForControl(oDoc2, sId))
                return (string)((IHTMLInputElement)oDoc2.all.item(sId, 0)).value;
            return null;
        }
        public static string[] RgsFromChlbx(bool fUse, CheckedListBox chlbx)
        {
            return RgsFromChlbx(fUse, chlbx, -1, false, null, false);
        }
        public static string[] RgsFromChlbxSport(bool fUse, CheckedListBox chlbx, string sSport, bool fMatch)
        {
            return RgsFromChlbx(fUse, chlbx, -1, false, sSport, fMatch);
        }
        public static string[] RgsFromChlbx(
            bool fUse,
            CheckedListBox chlbx,
            int iForceToggle,
            bool fForceOn,
            string sSport,
            bool fMatch)
        {
            string sSport2 = sSport == "Softball" ? "SB" : sSport;

            if (!fUse && sSport == null)
                return null;

            int c = chlbx.CheckedItems.Count;

            if (!fUse)
                c = chlbx.Items.Count;

            if (iForceToggle != -1)
                {
                if (fForceOn)
                    c++;
                else
                    c--;
                }

            string[] rgs = new string[c];
            int i = 0;

            if (!fUse)
                {
                int iT = 0;

                for (i = 0; i < c; i++)
                    {
                    rgs[iT] = (string) chlbx.Items[i];
                    if (sSport != null)
                        {
                        if ((rgs[iT].IndexOf(sSport) >= 0 && fMatch)
                            || (rgs[iT].IndexOf(sSport) == -1 && !fMatch)
                            || (rgs[iT].IndexOf(sSport2) >= 0 && fMatch)
                            || (rgs[iT].IndexOf(sSport2) == -1 && !fMatch))
                            {
                            iT++;
                            }
                        }
                    else
                        {
                        iT++;
                        }
                    }
                if (iT < c)
                    Array.Resize(ref rgs, iT);

                return rgs;
                }

            i = 0;
            foreach (int iChecked in chlbx.CheckedIndices)
                {
                if (iChecked == iForceToggle)
                    continue;
                rgs[i] = (string) chlbx.Items[iChecked];
                if (sSport != null)
                    {
                    if ((rgs[i].IndexOf(sSport) >= 0 && fMatch)
                        || (rgs[i].IndexOf(sSport) == -1 && !fMatch))
                        {
                        i++;
                        }
                    }
                else
                    {
                    i++;
                    }
                }
            if (fForceOn && iForceToggle != -1)
                rgs[i++] = (string) chlbx.Items[iForceToggle];

            if (i < c)
                Array.Resize(ref rgs, i);

            return rgs;
        }

        public static void UpdateChlbxFromRgs(CheckedListBox chlbx, string[] rgsSource, string[] rgsChecked, string[] rgsFilterPrefix, bool fCheckAll)
        {
            chlbx.Items.Clear();
            SortedList<string, int> mp = Utils.PlsUniqueFromRgs(rgsChecked);
			
            foreach (string s in rgsSource)
                {
                bool fSkip = false;

                if (rgsFilterPrefix != null)
                    {
                    fSkip = true;
                    foreach (string sPrefix in rgsFilterPrefix)
                        {
                        if (s.Length > sPrefix.Length && String.Compare(s.Substring(0, sPrefix.Length), sPrefix, true/*ignoreCase*/) == 0)
                            {
                            fSkip = false;
                            break;
                            }
                        }
                    }
                if (fSkip)
                    continue;

                CheckState cs;

                if (fCheckAll || mp.ContainsKey(s))
                    cs = CheckState.Checked;
                else
                    cs = CheckState.Unchecked;

                int i = chlbx.Items.Add(s, cs);
                }
        }
    }   
}