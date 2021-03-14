using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using mshtml;
using StatusBox;
using System.Runtime.InteropServices;

namespace Win32Win
{
    public class Win32
    {
        [DllImport("User32.dll")]
        public static extern IntPtr FindWindow(String lpClassName, String lpWindowName);

        [DllImport("User32.dll")]
        public static extern IntPtr SetForegroundWindow(IntPtr hWnd);

        [DllImport("User32.dll")]
        public static extern Boolean EnumChildWindows(IntPtr hWndParent, Delegate lpEnumFunc, IntPtr lParam);

        [DllImport("User32.dll")]
        public static extern Boolean EnumWindows(Delegate lpEnumFunc, IntPtr lParam);

        [DllImport("User32.dll")]
        public static extern IntPtr GetWindowText(IntPtr hWnd, StringBuilder s, IntPtr nMaxCount);

        [DllImport("User32.dll")]
        public static extern IntPtr GetWindowTextLength(IntPtr hwnd);

        [DllImport("user32.dll", EntryPoint = "GetDesktopWindow")]
        public static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll", EntryPoint = "PostMessage")]
        public static extern IntPtr PostMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", EntryPoint = "SendMessage")]
        public static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", EntryPoint = "SendMessage")]
        public static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, StringBuilder lParam);

        [DllImport("user32.dll", EntryPoint = "GetClassName")]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lParam, int cb);

        [DllImport("user32.dll", EntryPoint = "SetActiveWindow")]
        public static extern IntPtr SetActiveWindow(IntPtr hWnd);

        public const int BM_CLICK = 0x00F5;
        public const int WM_SETTEXT = 0x000C;
        public const int WM_GETTEXT = 0x000D;
        public const int WM_SETFOCUS = 0x0007;
        public const int WM_CHAR = 0x0102;

        public delegate int EnumChildWindowCallback(IntPtr hWnd, IntPtr lParam);
        public delegate int EnumWindowCallback(IntPtr hWnd, IntPtr lParam);

    }

    public class TrapFileDownload
    {
        private StatusRpt m_srpt;
        private string m_sName;
        private string m_sNameShort;
        private string m_sTarget;
        private string m_sChildFind;
        private string m_sAltChildFind;
        private AutoResetEvent m_evtCallerWaiting;
        private string m_sProgressWait;
        private string m_sClassNameFinding; // the class name for the child control we are looking for (this SHOULD be sent in through lParam, but I'm too tired to figure out how to coerce managed strings into lParams in native code and back again)

        public TrapFileDownload(StatusRpt srpt, string sExpectedName, string sExpectedNameShort, string sTarget, string sProgressWait, AutoResetEvent evt)
        {
            m_srpt = srpt;
            m_sName = sExpectedName;
            m_sNameShort = sExpectedNameShort;
            m_sTarget = sTarget;
            m_evtCallerWaiting = evt;
            m_sProgressWait = sProgressWait;
            m_srpt.LogData("Starting TaskTrap", 3, StatusRpt.MSGT.Body);
            Task taskTrap = new Task(TrapFileDownloadWork);

            taskTrap.Start();
        }

        private bool FWaitForWindowToBePresent(string sDlgClass, string sCaption, int n, string sErrorMsg,  out IntPtr hWnd)
        {
            while ((hWnd = Win32.FindWindow(sDlgClass, sCaption)) == IntPtr.Zero && n-- > 0)
            {
                Thread.Sleep(100);
                m_srpt.LogData(String.Format("FindWindow: 0x{0:X8}, n={1} ({2}/{3})", (Int64)hWnd, n, sDlgClass, sCaption), 3, StatusRpt.MSGT.Body);
            }

            m_srpt.LogData(String.Format("FindWindow DONE: 0x{0:X8}, n={1}", (Int64)hWnd, n), 3, StatusRpt.MSGT.Body);
            if (hWnd == IntPtr.Zero)
            {
                m_srpt.LogData(String.Format("{0}{1}, n={1}", sErrorMsg, hWnd, n), 3, StatusRpt.MSGT.Body);
                return false; // failed/timeout
            }
            return true;
        }

        private bool FSearchForExpectedTextInControls(string sTextToFind, string sAltTextToFind, string sClassName, string sPrologueMsg, IntPtr hWnd)
        {
            int cRetry = 3;

            while (cRetry >= 0)
                {
                if (FSearchForExpectedTextInControlsCore(sTextToFind, sAltTextToFind, sClassName, sPrologueMsg, hWnd))
                    return true;

                m_srpt.LogData(String.Format("Retry {0}...", cRetry), 3, StatusRpt.MSGT.Body);

                Thread.Sleep(100);
                cRetry--;
                }
            return false;
        }
        private bool FSearchForExpectedTextInControlsCore(string sTextToFind, string sAltTextToFind, string sClassName, string sPrologueMsg, IntPtr hWnd)
        {
            m_fFound = false;
            m_sChildFind = sTextToFind;
            m_sAltChildFind = sAltTextToFind;

            m_srpt.LogData(String.Format("{0} {1}/{2}", sPrologueMsg, sTextToFind, sAltTextToFind), 3, StatusRpt.MSGT.Body);
            m_sClassNameFinding = sClassName;
            Win32.EnumChildWindows(hWnd, new Win32.EnumChildWindowCallback(EnumChildCallback), IntPtr.Zero);
            if (!m_fFound)
            {
                m_srpt.LogData(String.Format("Couldn't find expected text: {0}/{1}", sTextToFind, sAltTextToFind), 3, StatusRpt.MSGT.Body);
                return false;
            }

            m_srpt.LogData(String.Format("Found expected text: {0}, {1}", m_sChildFound, m_fFound), 3, StatusRpt.MSGT.Body);
            return true;
        }

        string GetControlText(IntPtr hwnd)
        {
            StringBuilder sb = new StringBuilder((int)255 + 1);
            int cch;
            cch = (int)Win32.GetWindowText(hwnd, sb, (IntPtr)256);
            string s = sb.ToString().Trim();

            return s;
        }

        bool FIsWindowClassName(IntPtr hwnd, string sClass)
        {
            StringBuilder sb = new StringBuilder((int)255 + 1);
            Win32.GetClassName(hwnd, sb, 255);

            return String.Compare(sClass, sb.ToString(), true) == 0;
        }
        private bool FReplaceTextInControl(string sValidateText, string sReplaceText)
        {
            m_srpt.LogData(String.Format("Replacing {0} with {1} (handle = 0x{2:X8}", sValidateText, sReplaceText, (Int64)m_hwndFound), 3, StatusRpt.MSGT.Body);
            string sActual = GetControlText(m_hwndFound);
            m_srpt.LogData(String.Format("Actual text before: {0}", sActual), 3, StatusRpt.MSGT.Body);

            Win32.PostMessage(m_hwndFound, Win32.WM_SETFOCUS, IntPtr.Zero, IntPtr.Zero);
            m_srpt.LogData(String.Format("Posted message SETFOCUS"), 3, StatusRpt.MSGT.Body);

            foreach (char ch in sReplaceText)
                Win32.PostMessage(m_hwndFound, Win32.WM_CHAR, (IntPtr)ch, IntPtr.Zero);

            m_srpt.LogData(String.Format("Posted message WM_CHARs "), 3, StatusRpt.MSGT.Body);

            int cRetryForTypingLag = 30;
            while (cRetryForTypingLag-- > 0)
                {
                Thread.Sleep(10);

                sActual = GetControlText(m_hwndFound);
                if (String.Compare(sActual, sReplaceText, true) == 0)
                    break;
                }

            m_srpt.LogData(String.Format("Actual text after first: {0} [{1}msec]", sActual, 10 * (30 - cRetryForTypingLag)), 3, StatusRpt.MSGT.Body);
#if no
            Win32.PostMessage(m_hwndFound, Win32.WM_SETTEXT, IntPtr.Zero, Marshal.StringToCoTaskMemUni(sReplaceText));
            m_srpt.AddMessage(String.Format("Posted message SETTEXT"));

            sActual = GetControlText(m_hwndFound);
            m_srpt.AddMessage(String.Format("Actual text after first: {0}", sActual));

            Thread.Sleep(5000);
            Win32.PostMessage(m_hwndFound, Win32.WM_SETTEXT, IntPtr.Zero, Marshal.StringToHGlobalAnsi(sReplaceText));
            m_srpt.AddMessage(String.Format("Posted message SETTEXT"));

            sActual = GetControlText(m_hwndFound);
            m_srpt.AddMessage(String.Format("Actual text after second: {0}", sActual));
            Thread.Sleep(5000);
#endif //no
#if no
            Int32 cb = Win32.SendMessage(m_hwndFound, Win32.WM_GETTEXT, IntPtr.Zero, IntPtr.Zero).ToInt32();
            StringBuilder sb = new StringBuilder((int)cb + 1);
            Win32.SendMessage(m_hwndFound, Win32.WM_GETTEXT, (IntPtr)sb.Capacity, sb);

            m_srpt.AddMessage(String.Format("Verifying SETTEXT worked: Got {0}", sb.ToString()));
            if (String.Compare(sb.ToString(), sReplaceText, true) != 0)
                {
                m_srpt.AddMessage(String.Format("Verification FAILED!"));
                return false;
                }
#endif
            return true;
        }

        // This will sit around waiting for the right window to show up, and then handle it according to the
        // parameters given to us.
        // (designed to be run on a thread of its own)
        bool FHandleDialogAndClickButton(string sDlgClass, string sCaption, string sValidateText, string sValidateTextShort, string sValidateTextClassName, string sReplaceText, string sButtonToPress, bool fWaitForDialog)
        {
            IntPtr hWnd;
            int n = fWaitForDialog ? 1100 : 1;

            m_srpt.LogData(String.Format("FHandleDialogAndClickButton before FindWindow loop (Class={0}, Caption={1}) (WAIT_FOR_DIALOG={2}, VALIDATE_TEXT={3}, REPLACE_TEXT={4}, BUTTON_TO_PRESS={5})", sDlgClass, sCaption, fWaitForDialog, sValidateText ?? "null", sReplaceText ?? "null", sButtonToPress ?? "null"), 3, StatusRpt.MSGT.Body);

            if (!FWaitForWindowToBePresent(sDlgClass, sCaption, n, "TrapFileDownloadWork failed to find first window: ", out hWnd))
                return false;

            // now, enum the chilren to make sure that one of them has the text we are looking for!
            if (!FSearchForExpectedTextInControls(sValidateText, sValidateTextShort, sValidateTextClassName, "FHandleDialogAndClickButton before EnumChildWindows looking for", hWnd))
                return false;

            m_srpt.LogData($"About to try to replace text, {sValidateText} => {sReplaceText}", 3, StatusRpt.MSGT.Body);

            if (sReplaceText != null)
                {
                if (!FReplaceTextInControl(sValidateText, sReplaceText))
                    return false;
                }

            // ok yay, found it.  now we have to make like we clicked on a button...so let's find the button
            if (!FSearchForExpectedTextInControls(sButtonToPress, null, null, "FHandleDialogAndClickButton before EnumChildWindows looking for button text ", hWnd))
                return false;

            // now click on the button...
            m_srpt.LogData(String.Format("FHandleDialogAndClickButton SetActiveWindow({0})", hWnd), 3, StatusRpt.MSGT.Body);
            Win32.SetActiveWindow(hWnd);
            m_srpt.LogData(String.Format("FHandleDialogAndClickButton sending BM_CLICK to window {0}", m_hwndFound), 3, StatusRpt.MSGT.Body);
            Win32.SendMessage(m_hwndFound, Win32.BM_CLICK, IntPtr.Zero, IntPtr.Zero);

            Thread.Sleep(50);

            if (m_sProgressWait != null)
                {
                m_srpt.LogData(String.Format("Waiting for progress dialog to not be present: {0}", m_sProgressWait), 3, StatusRpt.MSGT.Body);
                n = 60;

                m_sChildFind = m_sProgressWait;
                do
                    {
                    m_fFound = false;
                    Win32.EnumWindows(new Win32.EnumWindowCallback(EnumWndCallback), IntPtr.Zero);
                    m_srpt.LogData(String.Format("EnumWindow found: {0} {1}", m_hwndFound, m_sChildFind), 3, StatusRpt.MSGT.Body);
                    Thread.Sleep(100);
                    } while (m_fFound && --n > 0);
                m_srpt.LogData(String.Format("Stopped waiting for progress dialog to disappear (countEnd: {0}, fFound: {1})", n, m_fFound), 3, StatusRpt.MSGT.Body);
                }
            return true;
        }


        /* T R A P  F I L E  D O W N L O A D  W O R K */
        /*----------------------------------------------------------------------------
        	%%Function: TrapFileDownloadWork
        	%%Qualified: Win32Win.TrapFileDownload.TrapFileDownloadWork
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public void TrapFileDownloadWork()
        {
            m_srpt.LogData("TrapFileDownloadWork top of loop", 3, StatusRpt.MSGT.Body);

            if (FHandleDialogAndClickButton("#32770", "Save As", m_sName, null, null, null, "&Save", true))
                {
                int c = 0;

                // do it again in case the first button click didn't work, go figure?
                m_srpt.LogData("TrapFileDownloadWork before repeat for click", 3, StatusRpt.MSGT.Body);
                while (FHandleDialogAndClickButton("#32770", "Save As", m_sName, null, null, null, "&Save", false))
                    {
                    m_srpt.LogData(String.Format("Had to click the button again ({0} times)...", ++c), 3, StatusRpt.MSGT.Body);
                    Thread.Sleep(700);
                    }

                m_srpt.LogData("TrapFileDownloadWork before SaveAs", 3, StatusRpt.MSGT.Body);
                FHandleDialogAndClickButton("#32770", "Save As", m_sName, m_sNameShort, "Edit", m_sTarget, "&Save", true);
                }

            if (m_evtCallerWaiting != null)
                m_evtCallerWaiting.Set();
        }

        private bool m_fFound;
        private string m_sChildFound;
        private IntPtr m_hwndFound;

        public int EnumWndCallback(IntPtr hWnd, IntPtr lParam)
        {
            StringBuilder sb = new StringBuilder(256);
            int cch;

            cch = (int)Win32.GetWindowText(hWnd, sb, (IntPtr)256);
            string s = sb.ToString().Trim();

            m_srpt.LogData(String.Format("EnumWndCallback: {0} =? {1} (hwnd = 0x{2:X8}", s, m_sChildFind, (Int64)hWnd), 3, StatusRpt.MSGT.Body);

            if (s.Contains(m_sChildFind) || (m_sAltChildFind != null && s.Contains(m_sAltChildFind)))
                {
                m_fFound = true;
                if (s.Contains(m_sChildFind))
                    m_sChildFound = m_sChildFind;
                else
                    m_sChildFound = m_sAltChildFind;
                m_hwndFound = hWnd;
                m_srpt.LogData("EnumWndCallback: FOUND", 3, StatusRpt.MSGT.Body);
                return 0;
                }
            return 1;
        }

        public int EnumChildCallback(IntPtr hWnd, IntPtr lParam)
        {
            StringBuilder sb = new StringBuilder(256);
            int cch;

            cch = (int)Win32.GetWindowText(hWnd, sb, (IntPtr)256);
            string s = sb.ToString().Trim();

            m_srpt.LogData(String.Format("EnumChildCallback: {0} =? {1} (hwnd = 0x{2:X8}", s, m_sChildFind, (Int64)hWnd), 3, StatusRpt.MSGT.Body);

            if (s == m_sChildFind)
                {
                if (m_sClassNameFinding == null || FIsWindowClassName(hWnd, m_sClassNameFinding))
                    {
                    m_srpt.LogData(String.Format("Found matching text, found matching classname"), 3, StatusRpt.MSGT.Body);
                    m_fFound = true;
                    m_hwndFound = hWnd;
                    m_srpt.LogData("EnumChildCallback: FOUND", 3, StatusRpt.MSGT.Body);
                    return 0;
                    }
                m_srpt.LogData(String.Format("Found matching text, CLASS NAME does not match"), 3, StatusRpt.MSGT.Body);
                }
            return 1;
        }

    }
#if no
    public class foo
    {
        private int hWnd;

        public delegate int Callback(int hWnd, int lParam);

        public int EnumChildGetValue(int hWnd, int lParam)
        {
            StringBuilder formDetails = new StringBuilder(256);
            int txtValue;
            string editText = "";
            txtValue = Win32.GetWindowText(hWnd, formDetails, 256);
            editText = formDetails.ToString().Trim();
            MessageBox.Show("Contains text of contro:" + editText);
            return 1;
        }

        public void foobar()
        {
            Callback myCallBack = new Callback(EnumChildGetValue);

            hWnd = Win32.FindWindow(null, "CallingWindow");
            if (hWnd == 0)
                {
                MessageBox.Show("Please Start Calling Window Application");
                }
            else
                {
                Win32.EnumChildWindows(hWnd, myCallBack, 0);
                }
        }
    }
#endif
}

    