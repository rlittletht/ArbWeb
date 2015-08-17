using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ArbWeb;
using mshtml;
using StatusBox;

namespace ArbWeb
{
    public partial class ArbWebControl2 : Form
    {
        private StatusRpt m_srpt;
        public ArbWebControl2()
        {
            InitializeComponent();
        }

        public WebBrowser AxWeb { get { return m_wbc; } }

        public IHTMLDocument  Document      {  get { return (IHTMLDocument)m_wbc.Document.DomDocument; } }
        public IHTMLDocument2 Document2     {  get { return (IHTMLDocument2)m_wbc.Document.DomDocument; } }
        public IHTMLDocument3 Document3     {  get { return (IHTMLDocument3)m_wbc.Document.DomDocument; } }
        public IHTMLDocument4 Document4     {  get { return (IHTMLDocument4)m_wbc.Document.DomDocument; } }
        public IHTMLDocument5 Document5     {  get { return (IHTMLDocument5)m_wbc.Document.DomDocument; } }

       public ArbWebControl2(StatusRpt srpt)
        {
#if notused
            m_plNewWindow3 = new List<DWebBrowserEvents2_NewWindow3EventHandler>();
            m_plBeforeNav2= new List<DWebBrowserEvents2_BeforeNavigate2EventHandler>();
#endif
            m_srpt = srpt;

            InitializeComponent();
        }

        bool m_fNavDone;
       public void ResetNav()
        {
            m_fNavDone = false;
        }


        void WaitForBrowserReady()
        {
            long s = 0;

            while (s < 200 && m_wbc.IsBusy)
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

        
        public void ReportNavState(string sTag)
        {
//            m_srpt.AddMessage(String.Format("{0}: Busy: {1}, State: {2}, m_fNavDone: {3}", sTag, m_axWebBrowser1.Busy, m_fNavDone, m_axWebBrowser1.ReadyState));
        }

       public bool FNavToPage(string sUrl)
        {
            ReportNavState("Entering FNavToPage: ");

            m_wbc.Stop();
            WaitForBrowserReady();
            m_fNavDone = false;
           m_wbc.Navigate(sUrl);
            m_wbc.Visible = true;

            return FWaitForNavFinish();
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
                if (m_wbc.ReadyState == WebBrowserReadyState.Complete) // m_fNavDone || 
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

        public void RefreshPage()
        {
            m_fNavDone = false;
            m_wbc.Refresh();
            FWaitForNavFinish();
        }
        private void TriggerDocumentDone(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            m_fNavDone = true;
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
        public bool FClickControl(IHTMLDocument2 oDoc2, string sId)
        {
//			m_srpt.AddMessage("Before clickcontrol: "+sId);
            ((IHTMLElement)(oDoc2.all.item(sId, 0))).click();
//			m_srpt.AddMessage("After clickcontrol");
            return FWaitForNavFinish();
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
