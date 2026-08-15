using System;
using System.Collections.Generic;
using System.Linq;

namespace ArbWeb;

public class CommonPrefixMap
{
#if nono
    private readonly Dictionary<string, List<string>> m_commonPrefixMap = new Dictionary<string, List<string>>();

    public int Count => m_commonPrefixMap.Count;

    public static int TokenCount(string s)
    {
        if (string.IsNullOrEmpty(s))
            return 0;

        return s.Split(' ').Length;
    }

    public CommonPrefixMap(List<string> sortedSites, int minTokens = 0)
    {
        string lastSiteFull = sortedSites[0];
        string currentCommonPrefix = lastSiteFull;
        List<string> currentSitesForPrefix = new List<string>();

        foreach (string site in sortedSites)
        {
            // see how much of the site name is common with the previous site name
            if (site == lastSiteFull && currentCommonPrefix != "")
            {
                // nothing we can infer here -- we're on the same site
                continue;
            }

            string commonPrefix = SiteShorter.GetCommonPrefix(site, lastSiteFull, minTokens);

            if (commonPrefix == "")
            {
                if (currentSitesForPrefix.Count > 0)
                    AddCommonPrefixSites(currentCommonPrefix, currentSitesForPrefix);

                currentSitesForPrefix.Clear();
                lastSiteFull = site;
                currentCommonPrefix = site;
                continue;
            }

            if (commonPrefix.Length < currentCommonPrefix.Length)
            {
                if (!currentCommonPrefix.StartsWith(commonPrefix))
                    throw new Exception($"Inconsistent common prefix: {currentCommonPrefix} vs {commonPrefix}");

                currentCommonPrefix = commonPrefix;
            }

            currentSitesForPrefix.Add(site);
        }

        if (currentSitesForPrefix.Count > 0)
            AddCommonPrefixSites(currentCommonPrefix, currentSitesForPrefix);
    }

    public void AddCommonPrefix(string prefix, string site)
    {
        if (!m_commonPrefixMap.ContainsKey(prefix))
            m_commonPrefixMap[prefix] = new List<string>();

        m_commonPrefixMap[prefix].Add(site);
    }

    public void AddCommonPrefixSites(string prefix, List<string> sites)
    {
        if (!m_commonPrefixMap.ContainsKey(prefix))
            m_commonPrefixMap[prefix] = new List<string>();

        m_commonPrefixMap[prefix].AddRange(sites);
    }

    public List<string> GetCommonPrefixes()
    {
        return m_commonPrefixMap.Keys.ToList();
    }

    public List<string> this[string prefix]
    {
        get
        {
            if (m_commonPrefixMap.ContainsKey(prefix))
                return m_commonPrefixMap[prefix];
            return new List<string>();
        }
    }

    public int Rank
    {
        get
        {

            return 0;
        }
    }
#endif // nono
}
