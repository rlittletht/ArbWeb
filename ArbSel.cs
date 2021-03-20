using System;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using StatusBox;
using System.Collections.Generic;
using System.Threading;
using HtmlAgilityPack;

namespace ArbWeb
{
	// Arbiter Selenium Control
	public class WebControl
	{
		private readonly IWebDriver m_driver;
		private IAppContext m_appContext;
		
		public IWebDriver Driver => m_driver;
		public string DownloadPath { get; set; }

		public WebControl(IAppContext context)
		{
			m_appContext = context;
			ChromeOptions options = new ChromeOptions();
			DownloadPath = $"{Environment.GetEnvironmentVariable("TEMP")}\\arb-{Guid.NewGuid().ToString()}";
			Directory.CreateDirectory(DownloadPath);
			
			options.AddUserProfilePreference("download.prompt_for_download", false);
			options.AddUserProfilePreference("download.default_directory", DownloadPath);

			m_driver = new ChromeDriver(options);
		}
		
		#region Page Navigation
		
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
			%%Function:WaitForPageLoad
			%%Qualified:ArbWeb.ArbWebControl_Selenium.WaitForPageLoad
		----------------------------------------------------------------------------*/
		public static void WaitForPageLoad(IAppContext appContext, IWebDriver driver, int maxWaitTimeInSeconds)
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
			appContext.StatusReport.LogData($"WaitForPageLoad elapsed: {timer.MsecFloat}", 1, StatusRpt.MSGT.Body);
		}

		public void WaitForPageLoad(int maxWaitTimeInSeconds = 500) => WaitForPageLoad(m_appContext, m_driver, maxWaitTimeInSeconds);

		public delegate bool WaitDelegate(IWebDriver driver);

		public void WaitForCondition(WaitDelegate waitDelegate, int msecTimeout = 500)
		{
			WebDriverWait wait = new WebDriverWait(m_driver, TimeSpan.FromMilliseconds(msecTimeout));

			wait.Until(
				(d) =>
				{
					try
					{
						return waitDelegate(d);
					}
					catch
					{
						return false;
					}
				});
		}
		
		#endregion

		#region Individual Control Interaction

		/*----------------------------------------------------------------------------
			%%Function:FCheckForControl
			%%Qualified:ArbWeb.ArbWebControl_Selenium.FCheckForControl
		----------------------------------------------------------------------------*/
		public static bool FCheckForControlId(IWebDriver driver, string sid)
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

		public bool FCheckForControlId(string sid) => FCheckForControlId(m_driver, sid);

		/*----------------------------------------------------------------------------
			%%Function:FClickControl
			%%Qualified:ArbWeb.ArbWebControl_Selenium.FClickControl
		----------------------------------------------------------------------------*/
		private static bool FClickControl(IAppContext appContext, IWebDriver driver, IWebElement element, string sidWaitFor = null)
		{
			element?.Click();

			if (sidWaitFor != null)
				return WaitForControl(driver, appContext, sidWaitFor);

			WaitForPageLoad(appContext, driver, 2000);
			return true;
		}
		
		/*----------------------------------------------------------------------------
			%%Function:FClickControl
			%%Qualified:ArbWeb.ArbWebControl_Selenium.FClickControl
		----------------------------------------------------------------------------*/
		public bool FClickControlName(string sName, string sidWaitFor = null)
		{
			m_appContext.StatusReport.LogData($"FClickControl {sName}", 5, StatusBox.StatusRpt.MSGT.Body);

			return FClickControl(m_appContext, m_driver, m_driver.FindElement(By.Name(sName)));
		}

		/*----------------------------------------------------------------------------
			%%Function:FClickControlId
			%%Qualified:ArbWeb.ArbWebControl_Selenium.FClickControlId
		----------------------------------------------------------------------------*/
		public bool FClickControlId(string sid, string sidWaitFor = null)
		{
			m_appContext.StatusReport.LogData($"FClickControl {sid}", 5, StatusBox.StatusRpt.MSGT.Body);

			return FClickControl(m_appContext, m_driver, m_driver.FindElement(By.Id(sid)));
		}

		/*----------------------------------------------------------------------------
			%%Function:FSetInputControlText
			%%Qualified:ArbWeb.ArbWebControl_Selenium.FSetInputControlText
		----------------------------------------------------------------------------*/
		public static bool FSetTextForInputControlName(IWebDriver driver, string sName, string sValue, bool fCheck)
		{
			IWebElement element = driver.FindElement(By.Name(sName));

			if (element == null)
				return false;

			string sOriginalValue = fCheck ? element.GetProperty("value") : null;

			element.Clear();
			element.SendKeys(sValue);


			if (fCheck)
				return String.Compare(sOriginalValue, sValue) != 0;

			return !fCheck;
		}

		public bool FSetTextForInputControlName(string sName, string sValue, bool fCheck) => FSetTextForInputControlName(m_driver, sName, sValue, fCheck);

		/*----------------------------------------------------------------------------
			%%Function:FSetCheckboxControlVal
			%%Qualified:ArbWeb.ArbWebControl_Selenium.FSetCheckboxControlIdVal
		----------------------------------------------------------------------------*/
		public static bool FSetCheckboxControlNameVal(IWebDriver driver, bool fChecked, string sName)
		{
			IWebElement element = driver.FindElement(By.Name(sName));
			string sValue = fChecked ? "true" : "false";
			string sActual = element.GetProperty("checked");

			if (String.Compare(element.GetProperty("checked"), sValue, true) != 0)
			{
				element.Click();
				return true;
			}

			return false;
		}

		public bool FSetCheckboxControlNameVal(bool fChecked, string sName) => FSetCheckboxControlNameVal(m_driver, fChecked, sName);

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

		public bool FSetCheckboxControlIdVal(bool fChecked, string sid) => FSetCheckboxControlIdVal(m_driver, fChecked, sid);

		/* S  G E T  C O N T R O L  V A L U E */
		/*----------------------------------------------------------------------------
        	%%Function: SGetControlValue
        	%%Qualified: ArbWeb.ArbWebControl.SGetControlValue
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
		public static string GetValueForControlId(IWebDriver driver, string sId)
		{
			IWebElement control = driver.FindElement(By.Id(sId));

			return control.GetProperty("value");
		}

		public string GetValueForControlId(string sId) => GetValueForControlId(m_driver, sId);

		public static bool FSetTextAreaTextForControlName(IWebDriver driver, string sName, string sValue, bool fCheck)
		{
			IWebElement element = driver.FindElement(By.Name(sName));

			string sOriginal = null;

			if (fCheck)
				sOriginal = element.GetAttribute("value");

			element.Clear();
			element.SendKeys(sValue);

			if (fCheck)
				return String.Compare(sValue, sOriginal, true) != 0;

			return true;
		}
		
		public bool FSetTextAreaTextForControlName(string sName, string sValue, bool fCheck) => FSetTextAreaTextForControlName(m_driver, sName, sValue, fCheck);

		#endregion

		#region Select/Option Interaction

		/*----------------------------------------------------------------------------
			%%Function:GetOptionValueFromFilterOptionText
			%%Qualified:ArbWeb.ArbWebControl_Selenium.GetOptionValueFromFilterOptionText
		----------------------------------------------------------------------------*/
		public string GetOptionValueFromFilterOptionTextForControlName(string sName, string sOptionText)
		{
			m_appContext.StatusReport.LogData($"SGetSelectIDFromDoc for id {sName}", 3, StatusRpt.MSGT.Body);

			string s = GetOptionValueForSelectControlNameOptionText(sName, sOptionText);

			m_appContext.StatusReport.LogData($"Return: {s}", 3, StatusRpt.MSGT.Body);
			return s;
		}

		/*----------------------------------------------------------------------------
			%%Function:GetOptionTextFromOptionValue
			%%Qualified:ArbWeb.ArbWebControl_Selenium.GetOptionTextFromOptionValue
		----------------------------------------------------------------------------*/
		public static string GetOptionTextFromOptionValueForControlName(IWebDriver driver, IAppContext context, string sName, string sOptionValue)
		{
			IWebElement selectElement = driver.FindElement(By.Name(sName));

			Dictionary<string, string> mpValueText = GetOptionsValueTextMappingFromControl(selectElement, context.StatusReport);
			if (mpValueText.ContainsKey(sOptionValue))
				return mpValueText[sOptionValue];

			return null;
		}

		public string GetOptionTextFromOptionValueForControlName(string sName, string sOptionName) => GetOptionTextFromOptionValueForControlName(m_driver, m_appContext, sName, sOptionName);

		/*----------------------------------------------------------------------------
			%%Function:FSetSelectControlText
			%%Qualified:ArbWeb.ArbWebControl_Selenium.FSetSelectControlText
		----------------------------------------------------------------------------*/
		public static bool FSetSelectedOptionTextForControlId(IWebDriver driver, IAppContext appContext, string sid, string sValue)
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

		public bool FSetSelectedOptionTextForControlId(string sid, string sValue) => FSetSelectedOptionTextForControlId(m_driver, m_appContext, sid, sValue);

		/*----------------------------------------------------------------------------
			%%Function:FSetSelectControlValue
			%%Qualified:ArbWeb.ArbWebControl_Selenium.FSetSelectControlValue
		----------------------------------------------------------------------------*/
		public static bool FSetSelectedOptionValueForControlName(IWebDriver driver, IAppContext appContext, string sName, string sValue)
		{
			appContext.StatusReport.LogData($"FSetSelectControlValue for name {sName}", 5, StatusRpt.MSGT.Body);

			SelectElement select = new SelectElement(driver.FindElement(By.Name(sName)));
			string sOriginal = select.SelectedOption.GetProperty("value");
			bool fChanged = false;

			if (String.Compare(sOriginal, sValue, true) != 0)
			{
				try
				{
					select.SelectByValue(sValue);
				}
				catch
				{
					return false;
				}
				fChanged = true;
			}

			appContext.StatusReport.LogData($"Return: {fChanged}", 5, StatusRpt.MSGT.Body);

			return fChanged;
		}

		public bool FSetSelectedOptionValueForControlName(string sName, string sValue) => FSetSelectedOptionValueForControlName(m_driver, m_appContext, sName, sValue);

		/*----------------------------------------------------------------------------
			%%Function:GetSelectedOptionTextFromSelectControl
			%%Qualified:ArbWeb.ArbWebControl_Selenium.GetSelectedOptionTextFromSelectControl
		----------------------------------------------------------------------------*/
		public static string GetSelectedOptionTextFromSelectControlName(IWebDriver driver, string sName)
		{
			IWebElement selectElement = driver.FindElement(By.Name(sName));
			SelectElement select = new SelectElement(selectElement);

			return select.SelectedOption?.Text;
		}

		public string GetSelectedOptionTextFromSelectControlName(string sName) => GetSelectedOptionTextFromSelectControlName(m_driver, sName);

		/*----------------------------------------------------------------------------
			%%Function:GetSelectedOptionValueFromSelectControlName
			%%Qualified:ArbWeb.ArbWebControl_Selenium.GetSelectedOptionValueFromSelectControlName
		----------------------------------------------------------------------------*/
		public static string GetSelectedOptionValueFromSelectControlName(IWebDriver driver, string sName)
		{
			IWebElement selectElement = driver.FindElement(By.Name(sName));
			SelectElement select = new SelectElement(selectElement);

			return select.SelectedOption.GetAttribute("value");
		}

		public string GetSelectedOptionValueFromSelectControlName(string sName) => GetSelectedOptionValueFromSelectControlName(m_driver, sName);

		/*----------------------------------------------------------------------------
			%%Function:GetOptionValueForSelectControlOptionText
			%%Qualified:ArbWeb.ArbWebControl_Selenium.GetOptionValueForSelectControlOptionText
		----------------------------------------------------------------------------*/
		public static string GetOptionValueForSelectControlNameOptionText(IWebDriver driver, string sName, string sOptionText)
		{
			IWebElement selectElement = driver.FindElement(By.Name(sName));
			Dictionary<string, string> mpValueText = GetOptionsValueTextMappingFromControl(selectElement, null);

			foreach (string sKey in mpValueText.Keys)
			{
				if (String.Compare(mpValueText[sKey], sOptionText, true) == 0)
					return sKey;
			}

			return null;
		}

		public string GetOptionValueForSelectControlNameOptionText(string sName, string sOptionText) => GetOptionValueForSelectControlNameOptionText(m_driver, sName, sOptionText);

		/*----------------------------------------------------------------------------
			%%Function:MpGetSelectValuesFromControl
			%%Qualified:ArbWeb.ArbWebControl_Selenium.MpGetSelectValuesFromControl
		----------------------------------------------------------------------------*/
		private static Dictionary<string, string> GetOptionsValueTextMappingFromControl(
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
					string sValue = option.GetAttributeValue("value", null);
					string sText = option.InnerText.Trim();
					
					if (mp.ContainsKey(sValue))
						srpt?.AddMessage(
							$"How strange!  Option '{sValue}' shows up more than once in the options list!",
							StatusRpt.MSGT.Warning);
					else
						mp.Add(sValue, sText);
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
		public static Dictionary<string, string> GetOptionsValueTextMappingFromControlId(IWebDriver driver, StatusRpt srpt, string sid)
		{
			MicroTimer timer = new MicroTimer();

			Dictionary<string, string> mp = GetOptionsValueTextMappingFromControl(driver.FindElement(By.Id(sid)), srpt);

			timer.Stop();
			srpt.LogData($"MpGetSelectValues({sid}) elapsed: {timer.MsecFloat}", 1, StatusRpt.MSGT.Body);
			return mp;
		}

		public Dictionary<string, string> GetOptionsValueTextMappingFromControlId(string sid) => GetOptionsValueTextMappingFromControlId(m_driver, m_appContext.StatusReport, sid);

		/*----------------------------------------------------------------------------
			%%Function:GetOptionsTextValueMappingFromControl
			%%Qualified:ArbWeb.ArbWebControl_Selenium.GetOptionsTextValueMappingFromControl
		----------------------------------------------------------------------------*/
		private static Dictionary<string, string> GetOptionsTextValueMappingFromControl(
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
					string sValue = option.GetAttributeValue("value", null);
					string sText = option.InnerText.Trim();

					if (mp.ContainsKey(sText))
						srpt?.AddMessage(
							$"How strange!  Option '{sText}' shows up more than once in the options list!",
							StatusRpt.MSGT.Warning);
					else
						mp.Add(sText, sValue);
				}
			}

			return mp;
		}

		/*----------------------------------------------------------------------------
			%%Function:GetOptionsTextValueMappingFromControlId
			%%Qualified:ArbWeb.ArbWebControl_Selenium.GetOptionsTextValueMappingFromControlId
		----------------------------------------------------------------------------*/
		public static Dictionary<string, string> GetOptionsTextValueMappingFromControlId(IWebDriver driver, StatusRpt srpt, string sid)
		{
			MicroTimer timer = new MicroTimer();

			Dictionary<string, string> mp = GetOptionsTextValueMappingFromControl(driver.FindElement(By.Id(sid)), srpt);

			timer.Stop();
			srpt.LogData($"GetOptionsTextValueMappingFromControlId({sid}) elapsed: {timer.MsecFloat}", 1, StatusRpt.MSGT.Body);
			return mp;
		}

		public Dictionary<string, string> GetOptionsTextValueMappingFromControlId(string sid) => GetOptionsTextValueMappingFromControlId(m_driver, m_appContext.StatusReport, sid);

		/*----------------------------------------------------------------------------
			%%Function:FResetMultiSelectOptions
			%%Qualified:ArbWeb.ArbWebControl_Selenium.FResetMultiSelectOptions

			Uncheck all of the items for this multiselect control
		----------------------------------------------------------------------------*/
		public static bool FResetMultiSelectOptionsForControlName(IWebDriver driver, string sName)
		{
			IWebElement selectElement = driver.FindElement(By.Name(sName));
			SelectElement select = new SelectElement(selectElement);

			select.DeselectAll();

			return true;
		}

		public bool FResetMultiSelectOptionsForControlName(string sName) => FResetMultiSelectOptionsForControlName(m_driver, sName);

		// if fValueIsValue == false, then sValue is the "text" of the option control
		/* F  S E L E C T  M U L T I  S E L E C T  O P T I O N */
		/*----------------------------------------------------------------------------
        	%%Function: FSelectMultiSelectOption
        	%%Qualified: ArbWeb.ArbWebControl.FSelectMultiSelectOption
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
		public static bool FSelectMultiSelectOptionValueForControlName(IWebDriver driver, string sName, string sValue)
		{
			IWebElement selectElement = driver.FindElement(By.Name(sName));
			SelectElement select = new SelectElement(selectElement);

			try
			{
				select.SelectByValue(sValue);
			}
			catch
			{
				return false;
			}

			return true;
		}

		public bool FSelectMultiSelectOptionValueForControlName(string sName, string sValue) => FSelectMultiSelectOptionValueForControlName(m_driver, sName, sValue);

		#endregion

		#region File Downloader
		public class FileDownloader
		{
			public delegate void StartDownload();
			
			private readonly string m_sExpectedFullName;
			private readonly string m_sTargetFile;
			private StartDownload m_startDownload;
			
			public FileDownloader(WebControl webControl, string expectedFile, string targetFile, StartDownload startDownload)
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
		#endregion
	}
}