using System;

namespace ArbWeb.Announcements;

public class Announcement
{
    public string Hint { get; set; }
    public string MatchString { get; set; }
    public bool ShowAssigners { get; set; }
    public bool ShowContacts { get; set; }
    public DateTime Date { get; set; }
    public string PostedBy { get; set; }
    public string Officials { get; set; }

    public string AnnouncementHtml { get; set; }
}
