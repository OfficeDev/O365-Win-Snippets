// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.OutlookServices;
using Microsoft.Office365.OutlookServices.Extensions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//Snippets in this file:
//
//GetOutlookClientAsync
//GetCalendarEventsAsync
//AddCalendarEventAsync
//AddCalendarEventWithArgsAsync
//UpdateCalendarEventAsync
//DeleteCalendarEventAsync

namespace O365_Win_Snippets
{
    public static class CalendarSnippets
    {
        private static OutlookServicesClient _outlookClient = null;

        /// <summary>
        /// Checks that an OutlookServicesClient object is available. 
        /// </summary>
        /// <returns>The OutlookServicesClient object. </returns>
        public static async Task<OutlookServicesClient> GetOutlookClientAsync()
        {
            if (_outlookClient != null && !String.IsNullOrEmpty(AuthenticationHelper.LastAuthority))
            {
                Debug.WriteLine("Got an Outlook client for Calendar.");
                return _outlookClient;
            }
            else
            {
                try
                {
                    //First, look for the authority used during the last authentication.
                    //If that value is not populated, use CommonAuthority.
                    string authority = null;

                    if (String.IsNullOrEmpty(AuthenticationHelper.LastAuthority))
                    {
                        authority = AuthenticationHelper.CommonAuthority;
                    }
                    else
                    {
                        authority = AuthenticationHelper.LastAuthority;
                    }

                    // Create an AuthenticationContext using this authority.
                    AuthenticationHelper._authenticationContext = new AuthenticationContext(authority);

                    // Set the value of _authenticationContext.UseCorporateNetwork to true so that you 
                    // can use this app inside a corporate intranet. If the value of UseCorporateNetwork 
                    // is true, you also need to add the Enterprise Authentication, Private Networks, and
                    // Shared User Certificates capabilities in the Package.appxmanifest file.
                    AuthenticationHelper._authenticationContext.UseCorporateNetwork = true;

                    //See the Discovery Service Sample (https://github.com/OfficeDev/Office365-Discovery-Service-Sample)
                    //for an approach that improves performance by storing the discovery service information in a cache.
                    DiscoveryClient discoveryClient = new DiscoveryClient(
                        async () => await AuthenticationHelper.GetTokenHelperAsync(AuthenticationHelper._authenticationContext, AuthenticationHelper.DiscoveryResourceId));

                    // Get the specified capability ("Calendar").
                    CapabilityDiscoveryResult result =
                        await discoveryClient.DiscoverCapabilityAsync("Calendar");

                    var token = await AuthenticationHelper.GetTokenHelperAsync(AuthenticationHelper._authenticationContext, result.ServiceResourceId);
                    // Check the token
                    if (String.IsNullOrEmpty(token))
                    {
                        // User cancelled sign-in
                        return null;
                    }
                    else
                    {

                        _outlookClient = new OutlookServicesClient(
                            result.ServiceEndpointUri,
                            async () => await AuthenticationHelper.GetTokenHelperAsync(AuthenticationHelper._authenticationContext, result.ServiceResourceId));
                        Debug.WriteLine("Got an Outlook client for Calendar.");
                        return _outlookClient;
                    }
                }
                // The following is a list of exceptions you should consider handling in your app.
                // In the case of this sample, the exceptions are handled by returning null upstream. 
                catch (DiscoveryFailedException dfe)
                {
                    Debug.WriteLine(dfe.Message);
                }
                catch (ArgumentException ae)
                {
                    Debug.WriteLine(ae.Message);
                }

                AuthenticationHelper._authenticationContext.TokenCache.Clear();

                return null;
            }
        }

        /// <summary>
        /// Gets a page of celendar events.
        /// </summary>
        /// <returns>A list of calendar events.</returns>
        public static async Task<List<IEvent>> GetCalendarEventsAsync()
        {
            try
            {
                // Make sure we have a reference to the Exchange client
                OutlookServicesClient client = await GetOutlookClientAsync();

                IPagedCollection<IEvent> eventsResults = await client.Me.Calendar.Events.ExecuteAsync();

                // You can access each event as follows.
                if (eventsResults.CurrentPage.Count > 0)
                {
                    string eventId = eventsResults.CurrentPage[0].Id;
                    Debug.WriteLine("First event:" + eventId);
                }

                return eventsResults.CurrentPage.ToList();
            }
            catch { return null; }
        }

        /// <summary>
        /// Adds a new event to user's default calendar
        /// </summary>
        public static async Task<string> AddCalendarEventAsync()
        {
            try
            {
                // Make sure we have a reference to the Exchange client
                var client = await GetOutlookClientAsync();

                Location location = new Location
                {
                    DisplayName = "Water cooler"
                };
                ItemBody body = new ItemBody
                {
                    Content = "Status updates, blocking issues, and next steps",
                    ContentType = BodyType.Text
                };

                Attendee[] attendees =  
            { 
                new Attendee  
                { 
                    Type = AttendeeType.Required, 
                    EmailAddress = new EmailAddress  
                    { 
                        Address = "mara@fabrikam.com" 
                    }, 
                }, 
            };

                Event newEvent = new Event
                {
                    Subject = "Weekly Sync",
                    Location = location,
                    Attendees = attendees,
                    Start = new DateTimeOffset(new DateTime(2014, 12, 1, 9, 30, 0)),
                    End = new DateTimeOffset(new DateTime(2014, 12, 1, 10, 0, 0)),
                    Body = body
                };

                await client.Me.Calendar.Events.AddEventAsync(newEvent);

                // Get the ID of the event. 
                string eventId = newEvent.Id;

                Debug.WriteLine("Added event: " + eventId);
                return eventId;
            }
            catch { return null; };


        }

        /// <summary>
        /// Adds a new event to user's default calendar
        /// </summary>
        /// <param name="LocationName">string. The name of the event location</param>
        /// <param name="BodyContent">string. The body of the event.</param>
        /// <param name="Attendees">string. semi-colon delimited list of invitee email addresses</param>
        /// <param name="EventName">string. The subject of the event</param>
        /// <param name="start">DateTimeOffset. The start date of the event</param>
        /// <param name="end">DateTimeOffset. The end date of the event</param>
        /// <param name="startTime">TimeSpan. The start hour:Min:Sec of the event</param>
        /// <param name="endTime">TimeSpan. The end hour:Min:Sec of the event</param>
        /// <returns>The Id of the event that was created; Otherwise, null.</returns>
        public static async Task<string> AddCalendarEventWithArgsAsync(
            string LocationName,
            string BodyContent,
            string Attendees,
            string EventName,
            DateTimeOffset start,
            DateTimeOffset end,
            TimeSpan startTime,
            TimeSpan endTime)
        {
            string newEventId = string.Empty;
            Location location = new Location();
            location.DisplayName = LocationName;
            ItemBody body = new ItemBody();
            body.Content = BodyContent;
            body.ContentType = BodyType.Text;
            string[] splitter = { ";" };
            var splitAttendeeString = Attendees.Split(splitter, StringSplitOptions.RemoveEmptyEntries);
            Attendee[] attendees = new Attendee[splitAttendeeString.Length];
            for (int i = 0; i < splitAttendeeString.Length; i++)
            {
                attendees[i] = new Attendee();
                attendees[i].Type = AttendeeType.Required;
                attendees[i].EmailAddress = new EmailAddress() { Address = splitAttendeeString[i] };
            }

            Event newEvent = new Event
            {
                Subject = EventName,
                Location = location,
                Attendees = attendees,
                Start = start,
                End = end,
                Body = body,
            };
            //Add new times to start and end dates.
            newEvent.Start = (DateTimeOffset?)CalcNewTime(newEvent.Start, start, startTime);
            newEvent.End = (DateTimeOffset?)CalcNewTime(newEvent.End, end, endTime);

            try
            {
                // Make sure we have a reference to the calendar client
                var exchangeClient = await GetOutlookClientAsync();

                // This results in a call to the service.
                await exchangeClient.Me.Events.AddEventAsync(newEvent);
                Debug.WriteLine("Added event: " + newEvent.Id);
                return newEvent.Id;
            }
            catch { return null; }
        }


        /// <summary>
        /// Updates an existing event in the user's default calendar
        /// </summary>
        /// <param name="eventId">string. The unique Id of the event to update</param>
        /// <param name="LocationName">string. The name of the event location</param>
        /// <param name="BodyContent">string. The body of the event.</param>
        /// <param name="Attendees">string. semi-colon delimited list of invitee email addresses</param>
        /// <param name="EventName">string. The subject of the event</param>
        /// <param name="start">DateTimeOffset. The start date of the event</param>
        /// <param name="end">DateTimeOffset. The end date of the event</param>
        /// <param name="startTime">TimeSpan. The start hour:Min:Sec of the event</param>
        /// <param name="endTime">TimeSpan. The end hour:Min:Sec of the event</param>
        /// <returns>IEvent. The updated event</returns>
        public static async Task<IEvent> UpdateCalendarEventAsync(string eventId,
            string LocationName,
            string BodyContent,
            string Attendees,
            string EventName,
            DateTimeOffset start,
            DateTimeOffset end,
            TimeSpan startTime,
            TimeSpan endTime)
        {
            // Make sure we have a reference to the Exchange client
            var client = await GetOutlookClientAsync();

            var eventToUpdate = await client.Me.Calendar.Events.GetById(eventId).ExecuteAsync();
            eventToUpdate.Attendees.Clear();
            string[] splitter = { ";" };
            var splitAttendeeString = Attendees.Split(splitter, StringSplitOptions.RemoveEmptyEntries);
            Attendee[] attendees = new Attendee[splitAttendeeString.Length];
            for (int i = 0; i < splitAttendeeString.Length; i++)
            {
                Attendee newAttendee = new Attendee();
                newAttendee.EmailAddress = new EmailAddress() { Address = splitAttendeeString[i] };
                newAttendee.Type = AttendeeType.Required;
                eventToUpdate.Attendees.Add(newAttendee);
            }

            eventToUpdate.Subject = EventName;
            Location location = new Location();
            location.DisplayName = LocationName;
            eventToUpdate.Location = location;
            eventToUpdate.Start = (DateTimeOffset?)CalcNewTime(eventToUpdate.Start, start, startTime);
            eventToUpdate.End = (DateTimeOffset?)CalcNewTime(eventToUpdate.End, end, endTime);
            ItemBody body = new ItemBody();
            body.ContentType = BodyType.Text;
            body.Content = BodyContent;
            eventToUpdate.Body = body;
            try
            {

                // Update the calendar event in Exchange
                await eventToUpdate.UpdateAsync();

                Debug.WriteLine("Updated event: " + eventToUpdate.Id);
                return eventToUpdate;

                // A note about Batch Updating
                // You can save multiple updates on the client and save them all at once (batch) by 
                // implementing the following pattern:
                // 1. Call UpdateAsync(true) for each event you want to update. Setting the parameter dontSave to true 
                //    means that the updates are registered locally on the client, but won't be posted to the server.
                // 2. Call exchangeClient.Context.SaveChangesAsync() to post all event updates you have saved locally  
                //    using the preceding UpdateAsync(true) call to the server, i.e., the user's Office 365 calendar.
            }
            catch { return null; }
        }

        /// <summary>
        /// Removes an event from the user's default calendar.
        /// </summary>
        /// <param name="eventId">string. The unique Id of the event to delete.</param>
        /// <returns></returns>
        public static async Task<IEvent> DeleteCalendarEventAsync(string eventId)
        {
            try
            {
                // Make sure we have a reference to the Exchange client
                var exchangeClient = await GetOutlookClientAsync();

                // Get the event to be removed from the Exchange service. This results in a call to the service.
                var eventToDelete = await exchangeClient.Me.Calendar.Events.GetById(eventId).ExecuteAsync();

                // Delete the event. This results in a call to the service.
                await eventToDelete.DeleteAsync(false);
                Debug.WriteLine("Deleted event: " + eventToDelete.Id);
                return eventToDelete;
            }
            catch { return null; }
        }

        //Helper method that creates a new DateTime for an updated meeting.

        /// <summary>
        /// Sets new time component of the datetimeoffset from the TimeSpan property of the UI
        /// </summary>
        /// <param name="OldDate">DateTimeOffset. The original date</param>
        /// <param name="NewDate">DateTimeOffset. The new date</param>
        /// <param name="newTime">TimeSpan. The new time</param>
        /// <returns>DateTimeOffset. The new start date/time</returns>
        private static DateTimeOffset CalcNewTime(DateTimeOffset? OldDate, DateTimeOffset NewDate, TimeSpan newTime)
        {
            //Default return value to New start date
            DateTimeOffset returnValue = NewDate;

            //Get original time components
            int hour = OldDate.Value.ToLocalTime().TimeOfDay.Hours;
            int min = OldDate.Value.ToLocalTime().TimeOfDay.Minutes;
            int second = OldDate.Value.ToLocalTime().TimeOfDay.Seconds;

            //Get new time components from TimeSpan updated by UI
            int newHour = newTime.Hours;
            int newMin = newTime.Minutes;
            int newSec = newTime.Seconds;

            //Update the new datetime by the difference between old and new time components
            returnValue = returnValue.AddHours(newHour - hour);
            returnValue = returnValue.AddMinutes(newMin - min);
            returnValue = returnValue.AddSeconds(newSec - second);

            return returnValue;
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