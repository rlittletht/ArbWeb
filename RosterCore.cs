using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using NUnit.Framework;

namespace ArbWeb
{
    public class RosterEntryPhones
    {
        public string m_sPhone1;
        public string m_sPhone2;
        public string m_sPhone3;

        public string CellPhone
        {
            get
            {
                if (m_sPhone1 != null && m_sPhone1.Contains("C:"))
                    return m_sPhone1;
                else if (m_sPhone2 != null && m_sPhone2.Contains("C:"))
                    return m_sPhone2;
                else if (m_sPhone3 != null && m_sPhone3.Contains("C:"))
                    return m_sPhone3;
                else
                    return m_sPhone1;
            }
        }

        public bool FHasPhoneNumber(int iPhone)
        {
            return !String.IsNullOrEmpty(NumberRawForPhone(iPhone));
        }

        static void ExtractNumberParts(string sNumberRaw, out string sNumber, out string sType)
        {
            sNumber = null;
            sType = null;

            if (String.IsNullOrEmpty(sNumberRaw))
                {
                return;
                }

            if (sNumberRaw.StartsWith("C:"))
                {
                sType = "Cellular";
                }
            else if (sNumberRaw.StartsWith("W:"))
                {
                sType = "Work";
                }
            else if (sNumberRaw.StartsWith("H"))
                {
                sType = "Home";
                }
            else
                {
                throw new Exception("Unknown phone type");
                }

            sNumber = sNumberRaw.Substring(2);
        }

        static string SConstructPhoneNumberFromParts(string sPhoneNumber, string sTypeDefault)
        {
            if (String.IsNullOrEmpty(sPhoneNumber))
                return sPhoneNumber;

            if (sPhoneNumber.StartsWith("C:") || sPhoneNumber.StartsWith("W:") || sPhoneNumber.StartsWith("H:"))
                return sPhoneNumber;

            return String.Format("{0}:{1}", sTypeDefault.Substring(0, 1), sPhoneNumber);
        }


        [TestCase("C:425-555-1212", "425-555-1212", "Cellular")]
        [TestCase("H:425-555-1212", "425-555-1212", "Home")]
        [TestCase("W:425-555-1212", "425-555-1212", "Work")]
        [TestCase(null, null, null)]
        [Test]
        public static void TestExtractNumberParts(string sNumberRaw, string sNumberExpected, string sTypeExpected)
        {
            string sNumberActual, sTypeActual;

            ExtractNumberParts(sNumberRaw, out sNumberActual, out sTypeActual);
            Assert.AreEqual(sNumberExpected, sNumberActual);
            Assert.AreEqual(sTypeExpected, sTypeActual);
        }
        string NumberRawForPhone(int iPhone)
        {
            string sNumberRaw = null;

            switch (iPhone)
                {
                case 1:
                    sNumberRaw = m_sPhone1;
                    break;
                case 2:
                    sNumberRaw = m_sPhone2;
                    break;
                case 3:
                    sNumberRaw = m_sPhone3;
                    break;
                }

            return sNumberRaw;
        }
        public void GetPhoneNumber(int iPhone, out string sNumber, out string sType)
        {
            string sNumberRaw = NumberRawForPhone(iPhone);
            ExtractNumberParts(sNumberRaw, out sNumber, out sType);
        }

        public void SetPhoneNumber(int iPhone, string sPhoneNumber, string sTypeDefault)
        {
            string sNumberRaw = SConstructPhoneNumberFromParts(sPhoneNumber, sTypeDefault);

            switch (iPhone)
                {
                case 1:
                    m_sPhone1 = sNumberRaw;
                    break;
                case 2:
                    m_sPhone2 = sNumberRaw;
                    break;
                case 3:
                    m_sPhone3 = sNumberRaw;
                    break;
                }
        }

        public void SetNextPhoneNumber(string sPhoneNumber, string sTypeDefault)
        {
            int iPhone = 1;

            if (String.IsNullOrEmpty(m_sPhone1))
                iPhone = 1;
            else if (String.IsNullOrEmpty(m_sPhone2))
                iPhone = 2;
            else if (String.IsNullOrEmpty(m_sPhone3))
                iPhone = 3;
            else
                return;

            SetPhoneNumber(iPhone, sPhoneNumber, sTypeDefault);
        }
    }
}
