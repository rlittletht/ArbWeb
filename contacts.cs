using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using NUnit.Framework;

namespace ArbWeb
{
    public class ContactEntry
    {
        public class TeamRelationship
        {
            private string m_sTeamName;
            private string m_sSport;
            private string m_sLevel;

            public string TeamName => m_sTeamName;
            public string Sport => m_sSport;
            public string Level => m_sLevel;

            public TeamRelationship(string sTeamName, string sSport, string sLevel)
            {
                m_sTeamName = sTeamName;
                m_sSport = sSport;
                m_sLevel = sLevel;
            }
        }

        string m_sEmail;
        string m_sFirst;
        string m_sLast;
        string m_sAddress1;
        string m_sAddress2;
        string m_sCity;
        string m_sState;
        string m_sZip;
        string m_sTitle;
        string m_sLastSignin;
        string m_sPhone1;
        string m_sPhone2;
        string m_sPhone3;

        public string Email { get { return m_sEmail; } set { m_sEmail = value; } } 
        public string First { get { return m_sFirst; } set { m_sFirst = value; } }
        public string Last { get { return m_sLast; } set { m_sLast = value; } }
        public string Address1 { get { return m_sAddress1; } set { m_sAddress1 = value; } }
        public string Address2 { get { return m_sAddress2; } set { m_sAddress2 = value; } }
        public string City { get { return m_sCity; } set { m_sCity = value; } }
        public string State { get { return m_sState; } set { m_sState = value; } }
        public string Zip { get { return m_sZip; } set { m_sZip = value; } }
        public string Title { get { return m_sTitle; } set { m_sTitle = value; } }
        public string LastSignin { get { return m_sLastSignin; } set { m_sLastSignin = value; } }
        public string Phone1 { get { return m_sPhone1; } set { m_sPhone1 = value; } }
        public string Phone2 { get { return m_sPhone2; } set { m_sPhone2 = value; } }
        public string Phone3 { get { return m_sPhone3; } set { m_sPhone3 = value; } }

        #region Phone Numbers

        #endregion
    }
}
