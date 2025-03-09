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
    public string InlineStylesheet { get; set; } = String.Empty;
    public int StyleClock { get; set; } = 0;
    public int Rank { get; set; } = -1;
    public bool NonArbweb { get; set; } = false;

    public bool Editable { get; set; }
    public bool Deletable { get; set; }

    public int PageIndex { get; set; } = -1;

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

    public string BuildAnnounceHtml(string id, string editHtml)
    {
        MatchString = @$"{WebAnnouncements.s_AnnounceDivIdPrefix}{id}";

        return @$"<div id=""{MatchString}"">{editHtml}</div>";
    }

    public static string AnnouncementIdFromMatchString(string matchString)
    {
        if (matchString == null)
            return null;

        return matchString.Substring(WebAnnouncements.s_AnnounceDivIdPrefix.Length);
    }

    public string AnnouncementID()
    {
        if (NonArbweb)
            return null;

        return AnnouncementIdFromMatchString(MatchString);
    }

    public Announcement()
    {}

    public Announcement(Announcement basedOn)
    {
        Hint = basedOn.Hint;
        MatchString = basedOn.MatchString;
        ShowAssigners = basedOn.ShowAssigners;
        ShowContacts = basedOn.ShowContacts;
        Date = basedOn.Date;
        Officials = basedOn.Officials;
        InlineStylesheet = basedOn.InlineStylesheet;
        StyleClock = basedOn.StyleClock;
        Rank = basedOn.Rank;
        Editable = basedOn.Editable;
        Deletable = basedOn.Deletable;
        PageIndex = basedOn.PageIndex;
        AnnouncementHtml = basedOn.AnnouncementHtml;
    }

    public bool Equals(Announcement announcement)
    {
        if (Officials != announcement.Officials)
            return false;
        if (ShowAssigners != announcement.ShowAssigners)
            return false;
        if (ShowContacts != announcement.ShowContacts)
            return false;
        if (AnnouncementHtml != announcement.AnnouncementHtml)
            return false;
        if (MatchString != announcement.MatchString)
            return false;
        return true;
    }

    public static Announcement CreateNonArbweb()
    {
        return 
            new Announcement
            {
                NonArbweb = true
            };
    }
}
