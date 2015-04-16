// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

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
            var outlookClient = await CalendarSnippets.GetOutlookClientAsync();
            return outlookClient != null;
        }

        public static async Task<bool> TryGetCalendarEventsAsync()
        {
            var events = await CalendarSnippets.GetCalendarEventsAsync();

            return events != null;
        }

        public static async Task<bool> TryCreateCalendarEventAsync()
        {
            var newEventId = await CalendarSnippets.AddCalendarEventAsync();

            if (newEventId == null)
                return false;

            //Cleanup
            await CalendarSnippets.DeleteCalendarEventAsync(newEventId);

            return true;
        }

        public static async Task<bool> TryCreateCalendarEventWithArgsAsync()
        {
            var newEventId = await CalendarSnippets.AddCalendarEventWithArgsAsync(
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
            await CalendarSnippets.DeleteCalendarEventAsync(newEventId);

            return true;
        }

        public static async Task<bool> TryUpdateCalendarEventAsync()
        {

            var newEventId = await CalendarSnippets.AddCalendarEventWithArgsAsync(
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
            var updatedEvent = await CalendarSnippets.UpdateCalendarEventAsync(newEventId,
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
            await CalendarSnippets.DeleteCalendarEventAsync(newEventId);


            return true;
        }

        public static async Task<bool> TryDeleteCalendarEventAsync()
        {
            var newEventId = await CalendarSnippets.AddCalendarEventAsync();

            if (newEventId == null)
                return false;

            // Delete the event
            var deletedEvent = await CalendarSnippets.DeleteCalendarEventAsync(newEventId);
            if (deletedEvent == null)
                return false;

            return true;
        }

    }
}

//********************************************************* 
// 
//O365-Win-Snippets, https://github.com/OfficeDev/O365-Win-Snippets
//
//Copyright (c) Microsoft Corporation
//All rights reserved. 
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// ""Software""), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:

// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.

// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
// 
//********************************************************* 