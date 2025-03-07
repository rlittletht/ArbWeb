using System;
using System.Text.RegularExpressions;

namespace ArbWeb.Announcements;

public class Announcement
{
    public string Hint { get; set; }
    public string MatchString { get; set; }
    public bool ShowAssigners { get; set; }
    public bool ShowContacts { get; set; }
    public DateTime Date { get; set; }
    public string Officials { get; set; }

    public bool Editable { get; set; }
    public bool Deletable { get; set; }

    public int PageIndex { get; set; }

    public string AnnouncementHtml { get; set; }

    public string ExtractEditHtml()
    {
        Regex divMatch = new Regex(@$"{MatchString}\w*""\w*>");

        Match match = divMatch.Match(AnnouncementHtml);

        if (match.Success == false)
            throw new Exception("Could not find announcement div in announcement html");

        string afterDiv = AnnouncementHtml.Substring(match.Index + match.Length);
        
        // now find the last div
        int lastDiv = afterDiv.LastIndexOf("</div>");

        return afterDiv.Substring(0, lastDiv);
    }

    public string RebuildAnnounceHtml(string id, string editHtml)
    {
        MatchString = @$"{WebAnnouncements.s_AnnounceDivIdPrefix}{id}";

        return @$"<div id=""{MatchString}"">{editHtml}</div>";
    }

    public string AnnouncementID()
    {
        return MatchString.Substring(WebAnnouncements.s_AnnounceDivIdPrefix.Length);
    }
}
