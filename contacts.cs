using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using NUnit.Framework;

namespace ArbWeb
{
    public class ContactEntry : RosterEntryNameAddress
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

        string m_sTitle;
        string m_sLastSignin;
        private List<TeamRelationship> m_plTeamRelationships;

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
            if (String.Compare(Address1, rste.Address1) != 0)
                return false;
            if (String.Compare(Address2, rste.Address2) != 0)
                return false;
            if (String.Compare(City, rste.City) != 0)
                return false;
            if (String.Compare(State, rste.State) != 0)
                return false;
            if (String.Compare(Zip, rste.Zip) != 0)
                return false;
            if (String.Compare(Title, rste.Title) != 0)
                return false;

            if (!FEqualsPhones(rste))
                return false;

            return true;
        }
    }

    public class ContactRoster
    {
        private List<ContactEntry> m_plce;

        public ContactRoster()
        {
            m_plce = new List<ContactEntry>();
        }

        public enum ReadState
        {
            ScanForHeader = 1,  // discard everything looking for header "Roster of Contacts"
            ScanForContact = 2, // discard everything scanning for a contact (value in first column)
            ReadingContact1 = 3, // Reading the first line of the contact (after scanning)
            ReadingContact2 = 4, // maybe reading a continuation of the contact line, OR start sites, OR start teams
            ReadingSites1 = 5, // definitely reading a site
            ReadingSites2 = 6, // maybe reading a continuation of a site, OR start teams, OR start contact
            ReadingTeams1 = 7, // definitely reading a team, then maybe team/site
            ReadingTeams2 = 8, // maybe reading a continuation of a team, OR start contact
        }

        ReadState RsScanForContact(string sLine, string[] rgsFields, ReadState rs)
        {
            // contacts will have 11 columns, and will have content in column 1
            if (rgsFields.Length < 11)
                return ReadState.ScanForContact;

            if (String.IsNullOrEmpty(rgsFields[0]))
                return ReadState.ScanForContact;

            return ReadState.ReadingContact1;
        }

        ReadState RsHandleReadingContact1(
            string sLine, 
            string[] rgsFields, 
            ReadState rs, 
            out string sName, 
            out string sAddressComplete,
            out string sPhone1,
            out string sPhone2,
            out string sPhone3,
            out string sEmail)
        {
            sName = rgsFields[0];
            sAddressComplete = rgsFields[1];
            sPhone1 = rgsFields[7];
            sPhone2 = rgsFields[8];
            sPhone3 = rgsFields[9];
            sEmail = rgsFields[rgsFields.Length - 1];

            return ReadState.ReadingContact2;
        }
#if contactparser
        public void ParseDownloadedRoster(string sFile)
        {
            TextReader tr = new StreamReader(sFile);
            string sLine;
            string[] rgsFields;
            ReadState rs = ReadState.ScanForContact;

            ContactEntry ceBuilding;
            int iColEmail = 11;

            while ((sLine = tr.ReadLine()) != null)
                {
                // first, change "foo, bar" into "foo bar" (get rid of quotes and the comma)
                // can't do that for this -- the address line is compound with many commas; can't lose that context
                // sLine = Regex.Replace(sLine, "\"([^\",]*),([^\",]*)\"", "$1$2");

                if (sLine.Length < iColEmail)
                    continue;

                rgsFields = Csv.LineToArray(sLine);

                if (rs == ReadState.ScanForHeader)
                    {
                    rs = RsScanForContact(sLine, rgsFields, rs);
                    // fallthrough. If we are still scanning, we won't do anything else...
                    }

                if (rs == ReadState.ReadingContact1)
                    {
                    ceBuilding = new ContactEntry();
                    rs = RsHandleReadingContact1(sLine, rgsFields, rs, ceBuilding);
                    }

                if (rs == ReadState.ReadingComments)
                    {
                    // when reading comments, we can get text in column 1; if the line ends with commas, then this is just
                    // a continuation of the comment (also be careful to look for another comment starting right after ours
                    // ends
                    if (FMatchGameCommentContinuation(sLine))
                        {
                        continue;
                        }

                    rs = ReadState.ReadingOfficials1;
                    // drop back to reading officials
                    }

                if (rs == ReadState.ReadingGame2)
                    rs = RsHandleReadingGame2(rgsFields, ref sGame, ref sDateTime, ref sLevel, ref sSite, ref sHome, ref sAway, rs);

                if (rs == ReadState.ReadingOfficials2)
                    rs = RsHandleReadingOfficials2(rst, rgsFields, mpNamePos, mpNameStatus, sNameLast, sPosLast, sStatusLast, rs);

                if (rs == ReadState.ReadingOfficials1)
                    rs = RsHandleReadingOfficials1(rst, fIncludeCanceled, sLine, rgsFields, mpNamePos, mpNameStatus, fCanceled, sSite, sGame,
                                                   sHome, sAway, sLevel, sSport, rs, ref sPosLast, ref sNameLast, ref sStatusLast, ref sDateTime, ref fOpenSlot, ref ump);

                if (FMatchGameArbiterFooter(sLine))
                    {
                    Debug.Assert(rs == ReadState.ReadingComments || rs == ReadState.ScanForHeader || rs == ReadState.ScanForGame,
                                 String.Format("Page break at illegal position: state = {0}", rs));
                    rs = ReadState.ScanForHeader;
                    continue;
                    }

                if (rs == ReadState.ScanForGame)
                    rs = RsHandleScanForGame(ref sGame, mpNamePos, mpNameStatus, sLine, ref sDateTime, ref sSport, ref sLevel, ref sSite,
                                             ref sHome, ref sAway, ref fCanceled, ref fIgnore, rs);

                if (rs == ReadState.ReadingGame1)
                    rs = RsHandleReadingGame1(ref sGame, rgsFields, ref sDateTime, ref sSport, ref sSite, ref sHome, ref sAway, rs);
                }

            return true;

        }
#endif 
    }
}
