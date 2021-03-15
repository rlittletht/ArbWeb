using System;
using System.Data;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using StatusBox;
using System.Collections.Generic;
using System.Threading;
using HtmlAgilityPack;
using OpenQA.Selenium.DevTools.V86.Input;

namespace ArbWeb
{
	// Arbiter Selenium Control
	public class ArbWebControl_Selenium
	{
		private readonly IWebDriver m_driver;
		private IAppContext m_appContext;
		
		public IWebDriver Driver => m_driver;
		public string DownloadPath { get; set; }

		public ArbWebControl_Selenium(IAppContext context)
		{
			m_appContext = context;
			ChromeOptions options = new ChromeOptions();
			DownloadPath = $"{Environment.GetEnvironmentVariable("TEMP")}\\arb-{Guid.NewGuid().ToString()}";
			Directory.CreateDirectory(DownloadPath);
			
			options.AddUserProfilePreference("download.prompt_for_download", false);
			options.AddUserProfilePreference("download.default_directory", DownloadPath);

			m_driver = new ChromeDriver(options);
		}
		
		/*----------------------------------------------------------------------------
			%%Function:FNavToPage
			%%Qualified:ArbWeb.ArbWebControl_Selenium.FNavToPage

		----------------------------------------------------------------------------*/
		public bool FNavToPage(string sUrl)
		{
			MicroTimer timer = new MicroTimer();
			
			try
			{
				m_driver.Navigate().GoToUrl(sUrl);
			}
			catch (Exception)
			{
				return false;
			}

			timer.Stop();
			m_appContext.StatusReport.LogData($"FNavToPage({sUrl}) elapsed: {timer.MsecFloat}", 1, StatusRpt.MSGT.Body);
			return true;
		}

		/*----------------------------------------------------------------------------
			%%Function:FCheckForControl
			%%Qualified:ArbWeb.ArbWebControl_Selenium.FCheckForControl
		----------------------------------------------------------------------------*/
		public static bool FCheckForControl(IWebDriver driver, string sid)
		{
			IWebElement element;

			try
			{
				element = driver.FindElement(By.Id(sid));
			}
			catch (OpenQA.Selenium.NoSuchElementException)
			{
				return false;
			}
			return element != null;
		}

		/*----------------------------------------------------------------------------
			%%Function:FSetInputControlText
			%%Qualified:ArbWeb.ArbWebControl_Selenium.FSetInputControlText
		----------------------------------------------------------------------------*/
		public static bool FSetInputControlText(IWebDriver driver, string sName, string sValue, bool fCheck)
		{
			IWebElement element = driver.FindElement(By.Name(sName));
			
			if (element == null)
				return false;

			string sOriginalValue = element.GetProperty("value");
			element.SendKeys(sValue);
			
			if (fCheck)
				return String.Compare(sOriginalValue, sValue) != 0;

			return false;
		}

		/*----------------------------------------------------------------------------
			%%Function:WaitForControl
			%%Qualified:ArbWeb.ArbWebControl_Selenium.WaitForControl

		----------------------------------------------------------------------------*/
		public static bool WaitForControl(IWebDriver driver, IAppContext appContext, string sid)
		{
			if (sid == null)
				return true;
			
			WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, 5));
			IWebElement element = wait.Until(theDriver => theDriver.FindElement(By.Id(sid)));

			return element != null;
		}
		
		/*----------------------------------------------------------------------------
			%%Function:FClickControl
			%%Qualified:ArbWeb.ArbWebControl_Selenium.FClickControl
		----------------------------------------------------------------------------*/
		public bool FClickControl(string sName, string sidWaitFor = null)
		{
			m_appContext.StatusReport.LogData(String.Format("FClickControl {0}", sName), 5, StatusBox.StatusRpt.MSGT.Body);
			
			IWebElement element = m_driver.FindElement(By.Name(sName));

			if (element != null)
				element.Click();

			//			m_srpt.AddMessage("After clickcontrol");
			if (sidWaitFor != null)
				return WaitForControl(m_driver, m_appContext, sidWaitFor);
			
			WaitForPageLoad(m_driver, 2000);
			return true;
		}

		/*----------------------------------------------------------------------------
			%%Function:FClickControlId
			%%Qualified:ArbWeb.ArbWebControl_Selenium.FClickControlId
		----------------------------------------------------------------------------*/
		public bool FClickControlId(string sid, string sidWaitFor = null)
		{
			m_appContext.StatusReport.LogData(String.Format("FClickControl {0}", sid), 5, StatusBox.StatusRpt.MSGT.Body);

			IWebElement element = m_driver.FindElement(By.Id(sid));

			if (element != null)
				element.Click();

			//			m_srpt.AddMessage("After clickcontrol");
			if (sidWaitFor != null)
				return WaitForControl(m_driver, m_appContext, sidWaitFor);

			WaitForPageLoad(m_driver, 2000);
			return true;
		}

		/*----------------------------------------------------------------------------
			%%Function:FSetCheckboxControlIdVal
			%%Qualified:ArbWeb.ArbWebControl_Selenium.FSetCheckboxControlIdVal
		----------------------------------------------------------------------------*/
		public static bool FSetCheckboxControlVal(IWebDriver driver, bool fChecked, string sName)
		{
			IWebElement element = driver.FindElement(By.Id(sName));
			string sValue = fChecked ? "true" : "false";
			string sActual = element.GetProperty("checked");

			if (String.Compare(element.GetProperty("checked"), sValue, true) != 0)
			{
				element.Click();
				return true;
			}

			return false;
		}

		public bool FSetCheckboxControlVal(bool fChecked, string sName)
		{
			return FSetCheckboxControlVal(m_driver, fChecked, sName);
		}
		
		/*----------------------------------------------------------------------------
			%%Function:FSetCheckboxControlIdVal
			%%Qualified:ArbWeb.ArbWebControl_Selenium.FSetCheckboxControlIdVal
		----------------------------------------------------------------------------*/
		public static bool FSetCheckboxControlIdVal(IWebDriver driver, bool fChecked, string sid)
		{
			IWebElement element = driver.FindElement(By.Id(sid));
			string sValue = fChecked ? "true" : "false";
			string sActual = element.GetProperty("checked");
			
			if (String.Compare(element.GetProperty("checked"), sValue, true) != 0)
			{
				element.Click();
				return true;
			}

			return false;
		}

		public bool FSetCheckboxControlIdVal(bool fChecked, string sid)
		{
			return FSetCheckboxControlIdVal(m_driver, fChecked, sid);
		}

		/* S  G E T  F I L T E R  I  D */
		/*----------------------------------------------------------------------------
        	%%Function: SGetFilterID
        	%%Qualified: ArbWeb.ArbWebControl.SGetFilterID
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
		public string SGetFilterID(string sName, string sValue)
		{
			m_appContext.StatusReport.LogData($"SGetSelectIDFromDoc for id {sName}", 3, StatusRpt.MSGT.Body);
			
			string s = SGetSelectIDFromDoc(sName, sValue);

			m_appContext.StatusReport.LogData($"Return: {s}", 3, StatusRpt.MSGT.Body);
			return s;
		}


		/* S  G E T  F I L T E R  I  D */
		/*----------------------------------------------------------------------------
        	%%Function: SGetFilterID
        	%%Qualified: ArbWeb.ArbWebControl.SGetFilterID
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
		public static string SGetSelectIDFromDoc(IWebDriver driver, IAppContext context, string sName, string sOptionName)
		{
			IWebElement selectElement = driver.FindElement(By.Name(sName));
			
			Dictionary<string, string> mpSelectValues = MpGetSelectValuesFromControl(selectElement, context.StatusReport);
			if (mpSelectValues.ContainsKey(sOptionName))
				return mpSelectValues[sOptionName];

			return null;
		}

		public string SGetSelectIDFromDoc(string sName, string sOptionName)
		{
			return SGetSelectIDFromDoc(m_driver, m_appContext, sName, sOptionName);
		}
		
		/*----------------------------------------------------------------------------
			%%Function:FSetSelectControlText
			%%Qualified:ArbWeb.ArbWebControl_Selenium.FSetSelectControlText
		----------------------------------------------------------------------------*/
		public static bool FSetSelectControlText(IWebDriver driver, IAppContext appContext, string sid, string sValue)
		{
			appContext.StatusReport.LogData($"FSetSelectControlText for id {sid}", 5, StatusRpt.MSGT.Body);
			
			SelectElement select = new SelectElement(driver.FindElement(By.Id(sid)));
			string sOriginal = select.SelectedOption.Text;
			bool fChanged = false;

			if (String.Compare(sOriginal, sValue, true) != 0)
			{
				select.SelectByText(sValue);
				fChanged = true;
			}
			
			appContext.StatusReport.LogData($"Return: {fChanged}", 5, StatusRpt.MSGT.Body);

			return fChanged;
		}
		
		public bool FSetSelectControlText(string sid, string sValue)
		{
			return FSetSelectControlText(m_driver, m_appContext, sid, sValue);
		}

		/*----------------------------------------------------------------------------
			%%Function:MpGetSelectValuesFromControl
			%%Qualified:ArbWeb.ArbWebControl_Selenium.MpGetSelectValuesFromControl
		----------------------------------------------------------------------------*/
		public static Dictionary<string, string> MpGetSelectValuesFromControl(
			IWebElement selectElement, 
			StatusRpt srpt)
		{
			string sHtml = selectElement.GetAttribute("innerHTML");
			
			HtmlDocument html = new HtmlDocument();
			html.LoadHtml(sHtml);

			HtmlNodeCollection options = html.DocumentNode.SelectNodes("//option");
			Dictionary<string, string> mp = new Dictionary<string, string>();

			if (options != null)
			{
				foreach (HtmlNode option in options)
				{
					if (mp.ContainsKey(option.InnerText))
						srpt.AddMessage(
							$"How strange!  Option '{option.InnerText}' shows up more than once in the options list!",
							StatusRpt.MSGT.Warning);
					else
						mp.Add(option.InnerText, option.SelectSingleNode("@value").InnerText);
				}
			}

			return mp;
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
		public static Dictionary<string, string> MpGetSelectValues(IWebDriver driver, StatusRpt srpt, string sid)
		{
			MicroTimer timer = new MicroTimer();
			
			Dictionary<string, string> mp = MpGetSelectValuesFromControl(driver.FindElement(By.Id(sid)), srpt);

			timer.Stop();
			srpt.LogData($"MpGetSelectValues({sid}) elapsed: {timer.MsecFloat}", 1, StatusRpt.MSGT.Body);
			return mp;
		}

		public Dictionary<string, string> MpGetSelectValues(string sid)
		{
			return MpGetSelectValues(m_driver, m_appContext.StatusReport, sid);
		}
		/*----------------------------------------------------------------------------
			%%Function:WaitForPageLoad
			%%Qualified:ArbWeb.ArbWebControl_Selenium.WaitForPageLoad
		----------------------------------------------------------------------------*/
		public void WaitForPageLoad(IWebDriver driver, int maxWaitTimeInSeconds)
		{
			MicroTimer timer = new MicroTimer();
			
			string state = string.Empty;
			try
			{
				WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(maxWaitTimeInSeconds));

				//Checks every 500 ms whether predicate returns true if returns exit otherwise keep trying till it returns ture
				wait.Until(d => {
					
					// use d instead of driver below?
					
					try
					{
						state = ((IJavaScriptExecutor)driver).ExecuteScript(@"return document.readyState").ToString();
					}
					catch (InvalidOperationException)
					{
						//Ignore
					}
					catch (NoSuchWindowException)
					{
						//when popup is closed, switch to last windows
						driver.SwitchTo().Window(driver.WindowHandles[driver.WindowHandles.Count - 1]);
					}
					//In IE7 there are chances we may get state as loaded instead of complete
					return state.Equals("complete", StringComparison.InvariantCultureIgnoreCase) 
					        || state.Equals("loaded", StringComparison.InvariantCultureIgnoreCase);

				});
			}
			catch (TimeoutException)
			{
				//sometimes Page remains in Interactive mode and never becomes Complete, then we can still try to access the controls
				if (!state.Equals("interactive", StringComparison.InvariantCultureIgnoreCase))
					throw;
			}
			catch (NullReferenceException)
			{
				//sometimes Page remains in Interactive mode and never becomes Complete, then we can still try to access the controls
				if (!state.Equals("interactive", StringComparison.InvariantCultureIgnoreCase))
					throw;
			}
			catch (WebDriverException)
			{
				if (driver.WindowHandles.Count == 1)
				{
					driver.SwitchTo().Window(driver.WindowHandles[0]);
				}
				state = ((IJavaScriptExecutor)driver).ExecuteScript(@"return document.readyState").ToString();
				if (!(state.Equals("complete", StringComparison.InvariantCultureIgnoreCase) || state.Equals("loaded", StringComparison.InvariantCultureIgnoreCase)))
					throw;
			}

			timer.Stop();
			m_appContext.StatusReport.LogData($"WaitForPageLoad elapsed: {timer.MsecFloat}", 1, StatusRpt.MSGT.Body);
		}

		public class FileDownloader
		{
			public delegate void StartDownload();
			
			private readonly string m_sExpectedFullName;
			private readonly string m_sTargetFile;
			private StartDownload m_startDownload;
			
			public FileDownloader(ArbWebControl_Selenium webControl, string expectedFile, string targetFile, StartDownload startDownload)
			{
				m_sExpectedFullName = Path.Combine(webControl.DownloadPath, expectedFile);
				if (targetFile == null)
				{
					m_sTargetFile = Path.Combine(
						webControl.DownloadPath,
						$"{System.Guid.NewGuid().ToString()}.{Path.GetExtension(expectedFile)}");
				}
				else
				{
					m_sTargetFile = targetFile;
				}
				
				// make sure that file doesn't already exist
				if (File.Exists(m_sExpectedFullName))
				{
					throw new Exception(
						$"File {m_sExpectedFullName} already exists! our temp download directory should start out empty! Someone not cleaning up?");
				}

				m_startDownload = startDownload;
			}
			
			public string GetDownloadedFile()
			{
				m_startDownload();
				
				// now wait for the file to be available and non-zero
				int cRetry = 100;
				while (--cRetry > 0)
				{
					Thread.Sleep(100);
					if (File.Exists(m_sExpectedFullName))
					{
						FileInfo info = new FileInfo(m_sExpectedFullName);

						if (info.Length > 0)
							break;
					}
				}

				if (cRetry <= 0)
					throw new Exception("file never downloaded?");

				File.Move(m_sExpectedFullName, m_sTargetFile);
				return m_sTargetFile;
			}
		}
		
	}
}