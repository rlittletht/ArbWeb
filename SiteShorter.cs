using ArbWeb.Games;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace ArbWeb;

public class SiteShorter
{
#if nono
    private readonly List<string> m_plsFullSiteNames;
    private readonly List<string> m_plsShortSiteNames;

    public SiteShorter()
    {
    }

    public SiteShorter(ScheduleGames games)
    {
        m_plsFullSiteNames = games.GetAllFullSiteNames();

        // now build the short site names for these sites
    }

    public List<string> GetSites()
    {
        return m_plsFullSiteNames;
    }

    public static string GetCommonPrefix(string s1, string s2, int minTokens = 0)
    {
        int minLength = Math.Min(s1.Length, s2.Length);
        int i = 0;
        int tokenCount = 1;
        while (i < minLength && s1[i] == s2[i])
        {
            i++;
            if (s1[i] == ' ' && i > 0 && s1[i - 1] != ' ')
                tokenCount++;
        }

        if (tokenCount < minTokens)
            return "";

        // chop trailing punctuation, spaces, and special characters
        string substring = s1.Substring(0, i);

        return substring.TrimEnd(new char[] { '#', '-', ':', ' ' });
    }

    [Test]
    [TestCase("Hartman #1", "Hartman #2", "Hartman")]
    [TestCase("Hartman 1", "Hartman 3", "Hartman")]
    public static void TestCommonPrefix(string s1, string s2, string sExpected)
    {
        string sCommon = GetCommonPrefix(s1, s2);
        Assert.AreEqual(sExpected, sCommon);
    }

    public class SiteReducer
    {
        private readonly CommonPrefixMap m_commonPrefixMap;

        public SiteReducer()
        {
        }

        public SiteReducer(List<string> sites)
        {
            List<string> sortedSites = new(sites);

            sortedSites.Sort();

            m_commonPrefixMap = new CommonPrefixMap(sortedSites);
        }

        public List<string> GetShortSites()
        {
            return m_commonPrefixMap.GetCommonPrefixes();
        }

        // this is clever but not needed
        public static string GetSortByLastTokenString(string s, int tokens)
        {
            // get the last 'tokens' tokens from the string s
            string[] parts = s.Split(' ');

            string suffix = string.Join(" ", parts.Skip(Math.Max(0, parts.Length - tokens)));
            string beforeSuffix = string.Join(" ", parts.Take(Math.Max(0, parts.Length - tokens)));

            return $"{suffix}|{beforeSuffix}";
        }

        [Test]
        [TestCase("Hartman #1", 1, "#1|Hartman")]
        [TestCase("Hartman #1", 2, "Hartman #1|")]
        [TestCase("Beaver Lake Park Consultants", 1, "Consultants|Beaver Lake Park")]
        [TestCase("Beaver Lake Park Consultants", 2, "Park Consultants|Beaver Lake")]
        [TestCase("Hartman Field 1", 1, "1|Hartman Field")]
        [TestCase("Hartman Field 1", 2, "Field 1|Hartman")]
        public static void TestGetSortByLastTokenString(string fullSite, int tokens, string expected)
        {
            Assert.AreEqual(expected, GetSortByLastTokenString(fullSite, tokens));
        }

        /*----------------------------------------------------------------------------
            %%Function: DifferentiateShortSites
            %%Qualified: ArbWeb.SiteShorter.SiteReducer.DifferentiateShortSites

            some common prefixes may be too short to be useful.
            For example, "Beaver Lake Park Consultants" and "Beaver Lake Park West #1"
            have a common prefix of "Beaver Lake Park". But really, there are two
            short sites -- Beaver Lake Park and Beaver Lake Park West.
        ----------------------------------------------------------------------------*/
        public CommonPrefixMap DifferentiateShortSites(CommonPrefixMap prefixMap)
        {
            CommonPrefixMap newPrefixMap = new CommonPrefixMap();

            foreach (string shortSite in prefixMap.GetCommonPrefixes())
            {
                List<string> sites = prefixMap[shortSite];
                List<string> sitesBy1MoreToken = new List<string>(sites);
                int currentTokenCount = CommonPrefixMap.TokenCount(shortSite);

                List<CommonPrefixMap> prefixMaps = new List<CommonPrefixMap>();

                // let's try extending by more tokens and find the best differentiation we can get
                int extraTokens = 1;

                while (true)
                {
                    CommonPrefixMap tokenMap = new CommonPrefixMap(sites, currentTokenCount + extraTokens);

                    if (tokenMap.Rank == 0)
                        break;

                    prefixMaps.Add(tokenMap);
                    extraTokens++;
                }

                // now find the best rank
            }
        }

        [Test]
        [TestCase(new string[] { "Hartman #1", "Hartman #2", "Hartman #3" }, new string[] { "Hartman" })]
        [TestCase(
            new string[]
            {
                "Beaver Lake Park Consultants", "Beaver Lake Park West #1", "Beaver Lake Park West #2", "Beaver Lake Park #1",
                "Beaver Lake Park #1"
            },
            new string[] { "Beaver Lake Park" })]
        public static void TestGetShortSites(string[] sites, string[] expectedShortSites)
        {
            SiteReducer reducer = new SiteReducer(sites.ToList());

            foreach (string shortSite in reducer.GetShortSites())
            {
                Assert.IsTrue(
                    expectedShortSites.Contains(shortSite, StringComparer.OrdinalIgnoreCase),
                    $"Short site '{shortSite}' not found in expected short sites.");
            }

            // and make sure that all expected short sites are in the shortSites list
            foreach (string expectedShortSite in expectedShortSites)
            {
                Assert.IsTrue(
                    reducer.GetShortSites().Contains(expectedShortSite, StringComparer.OrdinalIgnoreCase),
                    $"Expected short site '{expectedShortSite}' not found in short sites.");
            }
        }
    }

    public List<string> GetShortSites()
    {
        // sort the list
        List<string> sortedSites = new(m_plsFullSiteNames);

        sortedSites.Sort();


        // now let's try to find breaking points
        string lastSiteFull = sortedSites[0];
        string currentCommonPrefix = "";

        foreach (string site in sortedSites)
        {
            // see how much of the site name is common with the previous site name
            if (site == lastSiteFull)
            {
                // nothing we can infer here -- we're on the same site
                continue;
            }

            string commonPrefix = GetCommonPrefix(site, lastSiteFull);

            // find the common prefix between currentSiteFull and s
            if (currentCommonPrefix == "")
            {
                currentCommonPrefix = commonPrefix;
                continue;
            }
        }


        return m_plsShortSiteNames;
    }
#if no
    [Test]
    [TestCase(new string[] { "Hartman #1", "Hartman #2", "Hartman #3" }, new string[] { "Hartman" })]
    [TestCase(
        new string[] { "Beaver Lake Park Consultants", "Beaver Lake Park West #1", "Beaver Lake Park West #2", "Beaver Lake Park #1", "Beaver Lake Park #1" },
        new string[] { "Beaver Lake Park", "Beaver Lake Park West" })]
    [TestCase(
        new string[] { "Beaver Lake Park Consultants", "Beaver Lake Park Field 1", "Beaver Lake Park Field 2", "Beaver Lake Park West #1", "Beaver Lake Park West #2"},
        new string[] { "Beaver Lake Park", "Beaver Lake Park West" })]
    public static void Get(string[] fullSites, string[] expectedShortSites)
    {
        List<string> shortSites = SiteShorter.GetShortSites(fullSites);

        // verify shortSites matches the expected shortSites (case insensitive and order doesn't matter)
        foreach (string shortSite in shortSites)
        {
            if (shortSite != null)
            {
                Assert.IsTrue(
                    expectedShortSites.Contains(shortSite, StringComparer.OrdinalIgnoreCase),
                    $"Short site '{shortSite}' not found in expected short sites.");
            }

            // and make sure that all expected short sites are in the shortSites list
            foreach (string expectedShortSite in expectedShortSites)
            {
                if (expectedShortSite != null)
                {
                    Assert.IsTrue(
                        shortSites.Contains(expectedShortSite, StringComparer.OrdinalIgnoreCase),
                        $"Expected short site '{expectedShortSite}' not found in short sites.");
                }
            }
        }
#endif // no
#endif // nono
}
