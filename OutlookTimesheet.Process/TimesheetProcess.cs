using Microsoft.Office.Interop.Outlook;

namespace OutlookTimesheet.Process
{
    public class TimesheetProcess
    {
        public static IEnumerable<string> GetAllCalendarItems()
        {
            Application? oApp = null;
            NameSpace? mapiNamespace = null;
            MAPIFolder? CalendarFolder = null;
            Items? outlookCalendarItems = null;

            oApp = new Application();
            mapiNamespace = oApp.GetNamespace("MAPI");
            CalendarFolder = mapiNamespace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            outlookCalendarItems = CalendarFolder.Items;
            outlookCalendarItems.IncludeRecurrences = true;

            DateTime startDate = new DateTime(2022, 5, 9);
            DateTime endDate = new DateTime(2022, 5, 16);

            foreach (AppointmentItem item in outlookCalendarItems)
            {
                if (item.Start >= startDate && item.End <= endDate)
                {
                    if (item.IsRecurring)
                    {
                        Microsoft.Office.Interop.Outlook.RecurrencePattern rp = item.GetRecurrencePattern();
                        AppointmentItem? recur = null;

                        for (DateTime cur = startDate; cur <= endDate; cur = cur.AddDays(1))
                        {
                            string value = string.Empty;

                            try
                            {
                                recur = rp.GetOccurrence(cur);
                                value = item.Categories.First().ToString() + " -> " + recur.Subject + " -> " + cur.ToLongDateString() + recur;
                            }
                            catch
                            { }

                            if (!string.IsNullOrWhiteSpace(value))
                                yield return value;
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrWhiteSpace(item.Categories))
                            yield return item.Categories + " -> " + item.Subject + " -> " + item.Start.ToLongDateString();
                        else
                            yield return item.Subject + " -> " + item.Start.ToLongDateString();

                    }
                }
            }

        }
    }

}
