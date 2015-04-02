using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace O365_Win_Snippets
{
    public class CalendarStories
    {
        private static readonly string STORY_DATA_IDENTIFIER = Guid.NewGuid().ToString();

        public static async Task<bool> TryGetOutlookClientAsync()
        {
            var outlookClient = await CalendarOperations.GetOutlookClientAsync();
            return outlookClient != null;
        }

        public static async Task<bool> TryGetCalendarEventsAsync()
        {
            var events = await CalendarOperations.GetCalendarEventsAsync();

            return events != null;
        }

        public static async Task<bool> TryCreateCalendarEventAsync()
        {
            var newEventId = await CalendarOperations.AddCalendarEventAsync();

            if (newEventId == null)
                return false;

            //Cleanup
            await CalendarOperations.DeleteCalendarEventAsync(newEventId);

            return true;
        }

        public static async Task<bool> TryCreateCalendarEventWithArgsAsync()
        {
            var newEventId = await CalendarOperations.AddCalendarEventWithArgsAsync(
                            Guid.NewGuid().ToString(),
                            STORY_DATA_IDENTIFIER,
                            string.Empty,
                            Guid.NewGuid().ToString(),
                            new DateTimeOffset(DateTime.Now),
                            new DateTimeOffset(DateTime.Now),
                            new TimeSpan(DateTime.Now.Ticks),
                            new TimeSpan(DateTime.Now.Ticks)
                            );

            if (newEventId == null)
                return false;

            //Cleanup
            await CalendarOperations.DeleteCalendarEventAsync(newEventId);

            return true;
        }

        public static async Task<bool> TryUpdateCalendarEventAsync()
        {

            var newEventId = await CalendarOperations.AddCalendarEventWithArgsAsync(
                            "OrigLocationValue",
                            STORY_DATA_IDENTIFIER,
                            string.Empty,
                            Guid.NewGuid().ToString(),
                            new DateTimeOffset(DateTime.Now),
                            new DateTimeOffset(DateTime.Now),
                            new TimeSpan(DateTime.Now.Ticks),
                            new TimeSpan(DateTime.Now.Ticks)
                            );

            if (newEventId == null)
                return false;

            // Update our event
            var updatedEvent = await CalendarOperations.UpdateCalendarEventAsync(newEventId,
                           "NewLocationValue",
                           STORY_DATA_IDENTIFIER,
                           string.Empty,
                           Guid.NewGuid().ToString(),
                           new DateTimeOffset(DateTime.Now),
                           new DateTimeOffset(DateTime.Now),
                           new TimeSpan(DateTime.Now.Ticks),
                           new TimeSpan(DateTime.Now.Ticks)
                           );

            if (updatedEvent == null || updatedEvent.Id != newEventId)
                return false;

            //Cleanup
            await CalendarOperations.DeleteCalendarEventAsync(newEventId);


            return true;
        }

        public static async Task<bool> TryDeleteCalendarEventAsync()
        {
            var newEventId = await CalendarOperations.AddCalendarEventAsync();

            if (newEventId == null)
                return false;

            // Delete the event
            var deletedEvent = await CalendarOperations.DeleteCalendarEventAsync(newEventId);
            if (deletedEvent == null)
                return false;

            return true;
        }

    }
}
