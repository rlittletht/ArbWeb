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

        [DllImport("user32.dll", EntryPoint = "SendMessage")]
        public static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", EntryPoint = "SetActiveWindow")]
        public static extern IntPtr SetActiveWindow(IntPtr hWnd);

        public const int BM_CLICK = 0x00F5;
        public const int WM_SETTEXT = 0x000C;

        public delegate int EnumChildWindowCallback(IntPtr hWnd, IntPtr lParam);
        public delegate int EnumWindowCallback(IntPtr hWnd, IntPtr lParam);

    }

    public class TrapFileDownload
    {
        private StatusRpt m_srpt;
        private string m_sName;
        private string m_sTarget;
        private string m_sChildFind;
        private AutoResetEvent m_evtCallerWaiting;
        private string m_sProgressWait;

        public TrapFileDownload(StatusRpt srpt, string sExpectedName, string sTarget, string sProgressWait, AutoResetEvent evt)
        {
            m_srpt = srpt;
            m_sName = sExpectedName;
            m_sTarget = sTarget;
            m_evtCallerWaiting = evt;
            m_sProgressWait = sProgressWait;
            m_srpt.LogData("Starting TaskTrap", 3, StatusRpt.MSGT.Body);
            Task taskTrap = new Task(TrapFileDownloadWork);

            taskTrap.Start();
        }

        bool FHandleDialogAndClickButton(string sDlgClass, string sCaption, string sValidateText, string sReplaceText, string sButtonToPress, bool fWaitForDialog)
        {
            IntPtr hWnd;
            int n = fWaitForDialog ? 60 : 1;

            m_srpt.LogData(String.Format("FHandleDialogAndClickButton before FindWindow loop (Class={0}, Caption={1}) (WAIT_FOR_DIALOG={2})", sDlgClass, sCaption, fWaitForDialog), 3, StatusRpt.MSGT.Body);

            while ((hWnd = Win32.FindWindow(sDlgClass, sCaption)) == IntPtr.Zero && n-- > 0)
                {
                Thread.Sleep(1000);
                m_srpt.AddMessage(String.Format("FindWindow: {0}, n={1} ({2}/{3})", hWnd, n, sDlgClass, sCaption));
                }

            m_srpt.AddMessage(String.Format("FindWindow DONE: {0}, n={1}", hWnd, n));
            if (hWnd == IntPtr.Zero)
                {
                m_srpt.AddMessage(String.Format("TrapFileDownloadWork failed to find first window: {0}, n={1}", hWnd, n));
                return false; // failed/timeout
                }

            // now, enum the chilren to make sure that one of them has the text we are looking for!
            m_fFound = false;
            m_sChildFind = sValidateText;

            m_srpt.LogData(String.Format("FHandleDialogAndClickButton before EnumChildWindows looking for {0}", sValidateText), 3, StatusRpt.MSGT.Body);
            Win32.EnumChildWindows(hWnd, new Win32.EnumChildWindowCallback(EnumChildCallback), IntPtr.Zero);
            if (!m_fFound)
                {
                m_srpt.AddMessage(String.Format("Couldn't find expected text: {0}", sValidateText));
                return false;
                }

            m_srpt.AddMessage(String.Format("Found expected text: {0}, {1}", sValidateText, m_fFound));

            if (sReplaceText != null)
                {
                Win32.SendMessage(m_hwndFound, Win32.WM_SETTEXT, IntPtr.Zero, Marshal.StringToHGlobalAnsi(sReplaceText));
                //Thread.Sleep(5000);
                }

            // ok yay, found it.  now we have to make like we clicked on the open button...so let's find the open button
            m_fFound = false;
            m_sChildFind = sButtonToPress;
            Win32.EnumChildWindows(hWnd, new Win32.EnumChildWindowCallback(EnumChildCallback), IntPtr.Zero);
            if (!m_fFound)
                {
                m_srpt.AddMessage(String.Format("Couldn't find button to press: {0}", sButtonToPress));
                return false;
                }

            m_srpt.AddMessage(String.Format("Found expected button: {0}, {1}", sButtonToPress, m_hwndFound));
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
                    m_srpt.AddMessage(String.Format("EnumWindow found: {0} {1}", m_hwndFound, m_sChildFind));
                    Thread.Sleep(1000);
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

            if (FHandleDialogAndClickButton("#32770", "File Download", m_sName, null, "&Save", true))
                {
                int c = 0;

                // do it again in case the first button click didn't work, go figure?
                m_srpt.LogData("TrapFileDownloadWork before repeat for click", 3, StatusRpt.MSGT.Body);
                while (FHandleDialogAndClickButton("#32770", "File Download", m_sName, null, "&Save", false))
                    {
                    m_srpt.AddMessage(String.Format("Had to click the button again ({0} times)...", ++c));
                    Thread.Sleep(500);
                    }

                m_srpt.LogData("TrapFileDownloadWork before SaveAs", 3, StatusRpt.MSGT.Body);
                FHandleDialogAndClickButton("#32770", "Save As", m_sName, m_sTarget, "&Save", true);
                }

            if (m_evtCallerWaiting != null)
                m_evtCallerWaiting.Set();
        }

        private bool m_fFound;
        private IntPtr m_hwndFound;

        public int EnumWndCallback(IntPtr hWnd, IntPtr lParam)
        {
            StringBuilder sb = new StringBuilder(256);
            int cch;

            cch = (int)Win32.GetWindowText(hWnd, sb, (IntPtr)256);
            string s = sb.ToString().Trim();

            m_srpt.LogData(String.Format("EnumWndCallback: {0} =? {1}", s, m_sChildFind), 3, StatusRpt.MSGT.Body);

            if (s.Contains(m_sChildFind))
                {
                m_fFound = true;
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

            m_srpt.LogData(String.Format("EnumChildCallback: {0} =? {1}", s, m_sChildFind), 3, StatusRpt.MSGT.Body);

            if (s == m_sChildFind)
                {
                m_fFound = true;
                m_hwndFound = hWnd;
                m_srpt.LogData("EnumChildCallback: FOUND", 3, StatusRpt.MSGT.Body);
                return 0;
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

    