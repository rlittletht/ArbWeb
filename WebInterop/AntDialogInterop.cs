using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using TCore.StatusBox;
using TCore.WebControl;

namespace ArbWeb;

public class AntDialogInterop
{
    private readonly WebControl m_webControl;
    private readonly IStatusReporter m_statusReporter;

    public AntDialogInterop(WebControl webControl, IStatusReporter statusReporter)
    {
        m_webControl = webControl;
        m_statusReporter = statusReporter;
    }

    public static IWebElement WaitForDialog(IWebDriver driver, int msecToWait = 1000)
    {
            string xpathForDialogMask = "//div[@class='ant-modal-mask']";

            // wait for the dialog mask to be visible

            try
            {
                WebControl.WaitForCondition(
                    driver,
                    ExpectedConditions.ElementIsVisible(By.XPath(xpathForDialogMask)),
                    msecToWait);
            }
            catch (Exception e)
            {
                throw new Exception($"mask never became visible: {e}");
            }

            string xpathForDialogs = "//div[@class='ant-modal-wrap']";

            IReadOnlyCollection<IWebElement> dialogs = driver.FindElements(By.XPath(xpathForDialogs));

            // now find the dialog, if any, that is open
            foreach (IWebElement dialog in dialogs)
            {
                if (dialog.Displayed)
                    return dialog;
            }

            throw new Exception("No open dialog found");
    }

    /*----------------------------------------------------------------------------
        %%Function: WaitForDialog
        %%Qualified: ArbWeb.AntDialogInterop.WaitForDialog
    ----------------------------------------------------------------------------*/
    public IWebElement WaitForDialog(int msecToWait = 1000)
    {
        return WaitForDialog(m_webControl.Driver, msecToWait);
    }

    /*----------------------------------------------------------------------------
        %%Function: WaitForOneOfControlsOnActiveDialog
        %%Qualified: ArbWeb.AntDialogInterop.WaitForOneOfControlsOnActiveDialog
    ----------------------------------------------------------------------------*/
    public static IWebElement WaitForOneOfControlsOnActiveDialog(IWebDriver driver, IStatusReporter srpt, Func<IWebElement, IWebElement>[] findFuncs)
    {
        if (findFuncs == null || findFuncs.Length == 0)
            return null;

        WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, 5));
        IWebElement element = wait.Until(
            theDriver =>
            {
                // first, find the active dialog
                IWebElement dialog = WaitForDialog(theDriver);

                if (dialog == null)
                    return null;

                foreach (Func<IWebElement, IWebElement> findFunc in findFuncs)
                {
                    IWebElement element = findFunc(dialog);
                    if (element != null)
                        return element;
                }

                return null;
            });

        return element;
    }

    public IWebElement WaitForOneOfControlsOnActiveDialog(Func<IWebElement, IWebElement>[] findFuncs) => WaitForOneOfControlsOnActiveDialog(m_webControl.Driver, m_statusReporter, findFuncs);

    /*----------------------------------------------------------------------------
        %%Function: WaitForAndClickButtonByContent
        %%Qualified: ArbWeb.AntDialogInterop.WaitForAndClickButtonByContent
    ----------------------------------------------------------------------------*/
    public bool WaitForAndClickButtonByContent(IWebElement dialog, string content, int msecToWait = 1000)
    {
        string xpathForButton = $".//button[normalize-space(.)='{content}']";

        try
        {
            m_webControl.WaitForCondition(
                ExpectedConditions.ElementToBeClickable(By.XPath(xpathForButton)),
                msecToWait);
        }
        catch (Exception e)
        {
            throw new Exception($"button never became clickable: {e}");
        }

        IWebElement button = dialog.FindElement(By.XPath(xpathForButton));

        Utils.ThrowIfNot(button != null, $"no button found with content {content}");

        return WebControl.FClickControl(m_statusReporter, m_webControl.Driver, button);
    }
}
