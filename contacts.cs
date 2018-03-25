using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using NUnit.Framework;

namespace ArbWeb
{
    public class ContactEntry : RosterEntryPhones
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
        private List<TeamRelationship> m_plTeamRelationships;

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
        public List<TeamRelationship> TeamRelationships => m_plTeamRelationships;

        public ContactEntry()
        {
            m_plTeamRelationships = new List<TeamRelationship>();
        }
        public bool FEquals(ContactEntry rste)
        {
            if (String.Compare(First, rste.First) != 0)
                return false;
            if (String.Compare(Last, rste.Last) != 0)
                return false;
            if (String.Compare(Address1, rste.m_sAddress1) != 0)
                return false;
            if (String.Compare(Address2, rste.m_sAddress2) != 0)
                return false;
            if (String.Compare(City, rste.m_sCity) != 0)
                return false;
            if (String.Compare(State, rste.m_sState) != 0)
                return false;
            if (String.Compare(Zip, rste.m_sZip) != 0)
                return false;
            if (String.Compare(Title, rste.Title) != 0)
                return false;

            if (!FEqualsPhones(rste))
                return false;

            return true;
        }
    }
}
