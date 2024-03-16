using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using ArbWeb.Games;
using Microsoft.Identity.Client;

namespace ArbWeb.Reports
{
    public class CoverageReport
    {
        private IAppContext m_appContext;

        public CoverageReport()
        {
        } // for unit tests

        public CoverageReport(IAppContext appContext)
        {
            m_appContext = appContext;
        }

        class TimeSlotCoverage
        {
            private Dictionary<string, Dictionary<string, List<GameSlot>>> m_slotsBySite =
                new Dictionary<string, Dictionary<string, List<GameSlot>>>();

            public TimeSlotCoverage()
            {
            }

            public void AddSlot(GameSlot slot)
            {
                string site = slot.SiteShort;
                string subsite = slot.SubSite;

                if (!m_slotsBySite.ContainsKey(site))
                {
                    Dictionary<string, List<GameSlot>> newSlots = new Dictionary<string, List<GameSlot>>();
                    m_slotsBySite.Add(site, newSlots);
                }

                if (!m_slotsBySite[site].ContainsKey(subsite))
                    m_slotsBySite[site].Add(subsite, new List<GameSlot>());

                m_slotsBySite[site][subsite].Add(slot);
            }

            public IEnumerable<string> Sites => m_slotsBySite.Keys;

            public IEnumerable<string> SubSitesForSite(string site)
            {
                return m_slotsBySite[site].Keys;
            }

            public bool HasSite(string site) => m_slotsBySite.ContainsKey(site);

            public bool HasSubSite(string site, string subsite) => m_slotsBySite.ContainsKey(site) && m_slotsBySite[site].ContainsKey(subsite);

            public IEnumerable<GameSlot> Slots(string site, string subsite) => m_slotsBySite[site][subsite];
        }

        class Coverage
        {
            Dictionary<DateTime, Dictionary<TimeSpan, TimeSlotCoverage>> m_coverage =
                new Dictionary<DateTime, Dictionary<TimeSpan, TimeSlotCoverage>>();

            private Dictionary<string, string> m_siteRootMapping;

            public Coverage()
            {
            }

            public void AddSlot(GameSlot slot)
            {
                DateTime date = slot.Dttm.Date;
                TimeSpan time = slot.Dttm.TimeOfDay;

                if (!m_coverage.ContainsKey(date))
                    m_coverage.Add(date, new Dictionary<TimeSpan, TimeSlotCoverage>());

                if (!m_coverage[date].ContainsKey(time))
                    m_coverage[date].Add(time, new TimeSlotCoverage());

                m_coverage[date][time].AddSlot(slot);
            }

            private Dictionary<string, SortedList<string, string>> m_sites;

            static string KeyForSubSite(string site, string subsite)
            {
                if (subsite.ToUpper().Contains("CONSULTANT"))
                    return $"A-{subsite}";

                if (site.ToUpper().StartsWith(subsite.ToUpper()))
                    return $"A-{subsite}";

                // if there are more than one spaces in the name, and the final
                // space delimited substring matches the first, its also a consultant slot
                if (subsite.IndexOf(' ') != subsite.LastIndexOf(' '))
                {
                    string sSub = subsite.Substring(0, subsite.LastIndexOf(' '));

                    if (site.StartsWith(sSub))
                        return $"A-{subsite}";
                }

                return $"B-{subsite}";
            }

            void EnsureSiteInfo()
            {
                if (m_sites != null)
                    return;

                m_sites = new Dictionary<string, SortedList<string, string>>();

                foreach (Dictionary<TimeSpan, TimeSlotCoverage> timeCoverages in m_coverage.Values)
                {
                    foreach (TimeSlotCoverage timeSlotCoverage in timeCoverages.Values)
                    {
                        foreach (string site in timeSlotCoverage.Sites)
                        {
                            if (!m_sites.ContainsKey(site))
                                m_sites.Add(site, new SortedList<string, string>());

                            SortedList<string, string> subsites = m_sites[site];

                            foreach (string subsite in timeSlotCoverage.SubSitesForSite(site))
                            {
                                string subsiteKey = KeyForSubSite(site, subsite);

                                if (!subsites.ContainsKey(subsiteKey))
                                    subsites.Add(subsiteKey, subsite);
                            }
                        }
                    }
                }
            }

            public IEnumerable<DateTime> Dates => m_coverage.Keys;
            public IEnumerable<TimeSpan> TimeSlots(DateTime date) => m_coverage[date].Keys;

            public TimeSlotCoverage TimeSlotCoverage(DateTime date, TimeSpan time) => m_coverage[date][time];

            public int GetSitesCount()
            {
                EnsureSiteInfo();

                return m_sites.Keys.Count;
            }

            public IEnumerable<string> Sites
            {
                get
                {
                    EnsureSiteInfo();
                    return m_sites.Keys;
                }
            }

            public int GetSubsitesCountForSite(string site)
            {
                EnsureSiteInfo();

                return m_sites[site].Count;
            }

            public IEnumerable<string> SubSiteKeysForSite(string site)
            {
                EnsureSiteInfo();
                return m_sites[site].Keys;
            }

            public string SubsiteForSubsiteKey(string site, string key)
            {
                return m_sites[site][key];
            }

            public IEnumerable<string> SubSitesForSite(string site)
            {
                EnsureSiteInfo();
                return m_sites[site].Values;
            }


            public int GetTimeslotCountForDate(DateTime date)
            {
                return m_coverage[date].Keys.Count;
            }
        }


        void WriteHeader(StreamWriter sw)
        {
            sw.WriteLine(
                @"
<!DOCTYPE html>
<html>
  <head>
    <style>
      table.coverage { font-family: Calibri; font-size: 9pt; border-collapse: collapse; border; border: .25pt solid black }
      table.coverage td { padding-left: 3pt; padding-right: 3pt; vertical-align: top; }
      tr.heading td { text-align: center; vertical-align: bottom }
      col.left-bordered {border-left: 1.5pt solid black; border-top: 1.5pt solid black; border-bottom: 1.5pt solid black; }
      col.right-bordered {border-right: 1.5pt solid black; border-top: 1.5pt solid black; border-bottom: 1.5pt solid black; }
      col.inner-bordered {border-top: 1.5pt solid black; border-bottom: 1.5pt solid black; }
      tr.rowGroup1 td.shaded { background: rgb(228,228,228) }
    @page WordSection1
    	{size:11.0in 8.5in;
    	mso-page-orientation:landscape;
    	margin:.5in .5in .5in .5in;
    	mso-header-margin:.5in;
    	mso-footer-margin:.5in;
    	mso-paper-source:0;}
    div.WordSection1
    	{page:WordSection1;}
    </style>
    <meta http-equiv=""content-type"" content=""text/html; charset=UTF-8"">
    <title>coverage</title>
  </head>
  <body>
    <div class=WordSection1>
		");
        }

        void WriteCoverageHeaders(StreamWriter sw, Coverage coverage)
        {
            sw.WriteLine("<tr class='heading'>");
            sw.WriteLine("<td rowspan=2>Day</td>");
            sw.WriteLine("<td rowspan=2>Time</td>");

            foreach (string site in coverage.Sites)
            {
                sw.WriteLine($"<td colspan={coverage.GetSubsitesCountForSite(site)} style='{s_borderTop};{s_borderLeft};{s_borderRight}'>{site}</td>");
            }

            sw.WriteLine("</tr>");

            sw.WriteLine("<tr class='heading'>");
            StringBuilder sb = new StringBuilder();

            foreach (string site in coverage.Sites)
            {
                string subsiteLast = coverage.SubSitesForSite(site).Last();

                bool firstSite = true;

                foreach (string subsiteKey in coverage.SubSiteKeysForSite(site))
                {
                    string subsite = coverage.SubsiteForSubsiteKey(site, subsiteKey);

                    if (subsiteKey.StartsWith("A-"))
                        subsite = "Consultants";

                    sb.Clear();

                    sb.Append($"{s_borderBottom};");

                    if (firstSite)
                        sb.Append($"{s_borderLeft};");

                    firstSite = false;

                    if (subsiteLast == subsite)
                        sb.Append($"{s_borderRight};");

                    if (String.IsNullOrEmpty(subsite))
                        sw.WriteLine($"<td style='{sb.ToString()}'>&nbsp;</td>");
                    else
                        sw.WriteLine($"<td style='{sb.ToString()}'>{subsite}</td>");
                }
            }

            sw.WriteLine("</tr>");
        }


        private const string s_borderLeft = "border-left: 1pt solid black";
        private const string s_borderRight = "border-right: 1pt solid black";
        private const string s_borderTop = "border-top: 1pt solid black";
        private const string s_borderBottom = "border-bottom: 1pt solid black";
        private const string s_borderBottomThin = "border-bottom: 0.5pt solid black";
        private const string s_backGray = "background: rgb(232,232,232)";


        /*----------------------------------------------------------------------------
            %%Function: GenCoverageReport
            %%Qualified: ArbWeb.Reports.CoverageReport.GenCoverageReport

            Build a grid of coverage
        ----------------------------------------------------------------------------*/
        void GenCoverageReport(ScheduleGames games, string sReportFile, Roster rst, string[] rgsRoster, DateTime dttmStart, DateTime dttmEnd)
        {
            // we need to know who is working when...
            Coverage coverage = new Coverage();
            SortedList<string, int> plsSiteShort = Utils.PlsUniqueFromRgs(rgsRoster);

            SortedSet<string> plsSites = new SortedSet<string>();

            foreach (GameSlot gm in games.SortedGameSlotsByGameNumber)
            {
                if (!plsSites.Contains(gm.Site))
                    plsSites.Add(gm.Site);
            }

            Dictionary<string, string> mpSiteRoot = SiteRosterReport.MapCommonRootsFromList(plsSites);
            // coverage.AddSiteRootMapping(mpSiteRoot);

            // Map date => (time slot =>  list of fields; field => list of sub fields; field+sub-field => slots)
            foreach (GameSlot slot in games.SortedGameSlots)
            {
                if (slot.Dttm < dttmStart || slot.Dttm > dttmEnd)
                    continue;

                if (rgsRoster != null)
                    if (!plsSiteShort.ContainsKey(slot.SiteShort))
                        continue;

                coverage.AddSlot(slot);
            }

            // now, figure out how many fields we have

            using (StreamWriter sw = new StreamWriter(sReportFile, false, Encoding.Default))
            {
                WriteHeader(sw);
                sw.WriteLine("<table class='coverage'>");
                WriteCoverageHeaders(sw, coverage);
                bool shadeGroup = false;
                foreach (DateTime date in coverage.Dates)
                {
                    bool firstRow = true;
                    int slotCount = coverage.GetTimeslotCountForDate(date);
                    sw.WriteLine("<tr>");
                    sw.WriteLine($"<td rowspan={slotCount}>{date:dddd}</td>");
                    shadeGroup = !shadeGroup; // group items within 1 hour of previous slot, else flip

                    TimeSpan timeLast = new TimeSpan();

                    foreach (TimeSpan time in coverage.TimeSlots(date))
                    {
                        if (time.Hours - timeLast.Hours > 1 && !firstRow)
                            shadeGroup = !shadeGroup;

                        timeLast = time;

                        if (!firstRow)
                            sw.WriteLine("<tr>");

                        firstRow = false;
                        if (time.Hours < 12)
                        {
                            sw.WriteLine($"<td>{time.Hours}:{time.Minutes:00} AM</td>");
                        }
                        else
                        {
                            int hours = time.Hours > 12 ? time.Hours - 12 : time.Hours;

                            sw.WriteLine($"<td>{hours}:{time.Minutes:00} PM</td>");
                        }

                        TimeSlotCoverage timeSlotCoverage = coverage.TimeSlotCoverage(date, time);

                        // now interrogate the coverage using the sites we built the table with
                        foreach (string site in coverage.Sites)
                        {
                            string subsiteLast = coverage.SubSitesForSite(site).Last();
                            StringBuilder sb = new StringBuilder();

                            bool firstSite = true;

                            foreach (string subsite in coverage.SubSitesForSite(site))
                            {
                                sb.Clear();

                                sb.Append($"{s_borderBottomThin};");

                                if (firstSite)
                                    sb.Append($"{s_borderLeft};");

                                firstSite = false;

                                if (subsiteLast == subsite)
                                    sb.Append($"{s_borderRight};");

#if false
								// for now, we can't make shading work -- what if games are at 9/10/11/12/1/2/3/4/5 -- they are all within 1 hour!
								if (shadeGroup)
									sb.Append($"{s_backGray};");
#endif

                                if (timeSlotCoverage.HasSubSite(site, subsite))
                                {
                                    bool firstSlotRow = true;
                                    int openSlots = 0;

                                    sw.WriteLine($"<td style='{sb.ToString()}'>");

                                    foreach (GameSlot slot in timeSlotCoverage.Slots(site, subsite))
                                    {
                                        if (slot.Cancelled)
                                            continue;

                                        if (slot.Open && slot.Site.ToUpper().Contains("CONSULTANT"))
                                            continue;

                                        if (slot.Open)
                                        {
                                            openSlots++;
                                            continue;
                                        }

                                        RosterEntry rste = rst.RsteLookupEmail(slot.Email);

                                        if (!firstSlotRow)
                                            sw.WriteLine("<br/>");

                                        firstSlotRow = false;
                                        sw.Write($"{rste.NameShort}");
                                        if (slot.Status != "Accepted")
                                            sw.WriteLine("*");
                                        else
                                            sw.WriteLine();
                                    }

                                    while (openSlots-- > 0)
                                    {
                                        if (!firstSlotRow)
                                            sw.WriteLine("<br/>");

                                        firstSlotRow = false;
                                        sw.WriteLine("[]");
                                    }

                                    sw.WriteLine("</td>");
                                }
                                else
                                {
                                    sw.WriteLine($"<td style='{sb.ToString()}'>&nbsp;</td>");
                                }
                            }
                        }

                        sw.WriteLine("</tr>");
                    }
                }

                sw.WriteLine("</table>");
                sw.WriteLine("</body></div></html>");
            }
        }

        public void DoCoverageReport(CountsData gc, Roster rst, string[] rgsRoster, DateTime dttmStart, DateTime dttmEnd)
        {
            string sTempFile = $"{Environment.GetEnvironmentVariable("Temp")}\\temp{System.Guid.NewGuid().ToString()}.doc";

            GenCoverageReport(gc.Games, sTempFile, rst, rgsRoster, dttmStart, dttmEnd);
            // launch word with the file
            Process.Start(sTempFile);
            // System.IO.File.Delete(sTempFile);
        }
    }
}
