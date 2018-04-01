using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using NUnit.Framework;

namespace ArbWeb
{
    public class RosterEntryNameAddress : RosterEntryPhones
    {
        private string m_sEmail;
        private string m_sFirst;
        private string m_sLast;
        private string m_sAddress1;
        private string m_sAddress2;
        private string m_sCity;
        private string m_sState;
        private string m_sZip;

        public string Name { get { return String.Format("{0} {1}", First, Last); } }

        public string Email { get { return m_sEmail; } set { m_sEmail = value; } }
        public string First { get { return m_sFirst; } set { m_sFirst = value; } }
        public string Last { get { return m_sLast; } set { m_sLast = value; } }
        public string Address1 { get { return m_sAddress1; } set { m_sAddress1 = value; } }
        public string Address2 { get { return m_sAddress2; } set { m_sAddress2 = value; } }
        public string City { get { return m_sCity; } set { m_sCity = value; } }
        public string State { get { return m_sState; } set { m_sState = value; } }
        public string Zip { get { return m_sZip; } set { m_sZip = value; } }

        /* S E T  E M A I L */
        /*----------------------------------------------------------------------------
			%%Function: SetEmail
			%%Qualified: ArbWeb.RSTE.SetEmail
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
        public void SetEmail(string s)
        {
            Email = s.Substring(s.IndexOf(":") + 1);
        }
    }
    public class RosterEntryPhones
    {
        private string m_sPhone1;
        private string m_sPhone2;
        private string m_sPhone3;

        public string Phone1 { get { return m_sPhone1; } set { m_sPhone1 = value; } }
        public string Phone2 { get { return m_sPhone2; } set { m_sPhone2 = value; } }
        public string Phone3 { get { return m_sPhone3; } set { m_sPhone3 = value; } }

        public string CellPhone
        {
            get
            {
                if (Phone1 != null && Phone1.Contains("C:"))
                    return Phone1;
                else if (Phone2 != null && Phone2.Contains("C:"))
                    return Phone2;
                else if (Phone3 != null && Phone3.Contains("C:"))
                    return Phone3;
                else
                    return Phone1;
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
                    sNumberRaw = Phone1;
                    break;
                case 2:
                    sNumberRaw = Phone2;
                    break;
                case 3:
                    sNumberRaw = Phone3;
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
                    Phone1 = sNumberRaw;
                    break;
                case 2:
                    Phone2 = sNumberRaw;
                    break;
                case 3:
                    Phone3 = sNumberRaw;
                    break;
                }
        }

        public void SetNextPhoneNumber(string sPhoneNumber, string sTypeDefault)
        {
            int iPhone = 1;

            if (String.IsNullOrEmpty(Phone1))
                iPhone = 1;
            else if (String.IsNullOrEmpty(Phone2))
                iPhone = 2;
            else if (String.IsNullOrEmpty(Phone3))
                iPhone = 3;
            else
                return;

            SetPhoneNumber(iPhone, sPhoneNumber, sTypeDefault);
        }


        public bool FEqualsPhones(RosterEntryPhones rstep)
        {
            if (!(String.IsNullOrEmpty(Phone1) && String.IsNullOrEmpty(rstep.Phone1)) && String.Compare(Phone1, rstep.Phone1) != 0)
                return false;
            if (!(String.IsNullOrEmpty(Phone2) && String.IsNullOrEmpty(rstep.Phone2)) && String.Compare(Phone2, rstep.Phone2) != 0)
                return false;
            if (!(String.IsNullOrEmpty(Phone3) && String.IsNullOrEmpty(rstep.Phone3)) && String.Compare(Phone3, rstep.Phone3) != 0)
                return false;

            return true;
        }
    }

    public interface IRosterEntry
    {
        void SetEmail(string sEmail);
        string Email { get; set; }
        bool FEqualsMisc(IRosterEntry irste);
        bool Marked { get; set; }
    }

    public interface IRoster
    {
        List<string> PlsMiscLookupEmail(string sEmail);
        bool IsQuick { get; }
        bool IsUploadableQuickroster { get; }
        void Add(IRosterEntry rste);
        IRosterEntry IrsteLookupEmail(string sEmail);
        List<IRosterEntry> PlirsteUnmarked();
        List<IRosterEntry> Plirste { get; }
        void WriteRoster(string sOutFile);
        IRosterEntry CreateRosterEntry();
    }
}
