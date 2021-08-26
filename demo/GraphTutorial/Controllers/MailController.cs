// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using GraphTutorial.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Web;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using TimeZoneConverter;
using System.Collections.ObjectModel;

namespace GraphTutorial.Controllers
{
    public class MailController : Controller
    {
        private readonly GraphServiceClient _graphClient;
        private readonly ILogger<HomeController> _logger;

        public MailController(
            GraphServiceClient graphClient,
            ILogger<HomeController> logger)
        {
            _graphClient = graphClient;
            _logger = logger;
        }

        // <IndexSnippet>
        // Minimum permission scope needed for this view
        [AuthorizeForScopes(Scopes = new[] { "Calendars.Read" })]
        public async Task<IActionResult> Index()
        {
            try
            {
                var userTimeZone = TZConvert.GetTimeZoneInfo(
                    User.GetUserGraphTimeZone());
                var startOfWeekUtc = MailController.GetUtcStartOfWeekInTimeZone(
                    DateTime.Today, userTimeZone);

                ObservableCollection<GraphSDKDemo.Models.Message> events = await GetUserWeekCalendar(startOfWeekUtc);
                
                var model = new MailViewModel(events);

                return View(model);
            }
            catch (ServiceException ex)
            {
                if (ex.InnerException is MicrosoftIdentityWebChallengeUserException)
                {
                    throw;
                }

                return View(new CalendarViewModel())
                    .WithError("Error getting calendar view", ex.Message);
            }
        }
        // </IndexSnippet>

        // <CalendarNewGetSnippet>
        // Minimum permission scope needed for this view
        [AuthorizeForScopes(Scopes = new[] { "Calendars.ReadWrite" })]
        public IActionResult New()
        {
            return View();
        }
        // </CalendarNewGetSnippet>

        // <CalendarNewPostSnippet>
        [HttpPost]
        [ValidateAntiForgeryToken]
        [AuthorizeForScopes(Scopes = new[] { "Calendars.ReadWrite" })]
        public async Task<IActionResult> New([Bind("Subject,Attendees,Start,End,Body")] NewEvent newEvent)
        {
            var timeZone = User.GetUserGraphTimeZone();

            // Create a Graph event with the required fields
            var graphEvent = new Event
            {
                Subject = newEvent.Subject,
                Start = new DateTimeTimeZone
                {
                    DateTime = newEvent.Start.ToString("o"),
                    // Use the user's time zone
                    TimeZone = timeZone
                },
                End = new DateTimeTimeZone
                {
                    DateTime = newEvent.End.ToString("o"),
                    // Use the user's time zone
                    TimeZone = timeZone
                }
            };

            // Add body if present
            if (!string.IsNullOrEmpty(newEvent.Body))
            {
                graphEvent.Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = newEvent.Body
                };
            }

            // Add attendees if present
            if (!string.IsNullOrEmpty(newEvent.Attendees))
            {
                var attendees =
                    newEvent.Attendees.Split(';', StringSplitOptions.RemoveEmptyEntries);

                if (attendees.Length > 0)
                {
                    var attendeeList = new List<Attendee>();
                    foreach (var attendee in attendees)
                    {
                        attendeeList.Add(new Attendee{
                            EmailAddress = new EmailAddress
                            {
                                Address = attendee
                            },
                            Type = AttendeeType.Required
                        });
                    }

                    graphEvent.Attendees = attendeeList;
                }
            }

            try
            {
                // Add the event
                await _graphClient.Me.Events
                    .Request()
                    .AddAsync(graphEvent);

                // Redirect to the calendar view with a success message
                return RedirectToAction("Index").WithSuccess("Event created");
            }
            catch (ServiceException ex)
            {
                // Redirect to the calendar view with an error message
                return RedirectToAction("Index")
                    .WithError("Error creating event", ex.Error.Message);
            }
        }
        // </CalendarNewPostSnippet>

        // <GetCalendarViewSnippet>
        private async Task<ObservableCollection<GraphSDKDemo.Models.Message>> GetUserWeekCalendar(DateTime startOfWeekUtc)
        {
            // Get all messages, whether or not they are in the Inbox
            var userMessages = await _graphClient.Me.Messages.Request().Top(10)
                                            .Select("sender, from, subject, importance")
                                            .GetAsync();
            var MyMessages = new ObservableCollection<GraphSDKDemo.Models.Message>();
            foreach (var message in userMessages)
            {
                MyMessages.Add(new GraphSDKDemo.Models.Message
                {
                    Id = message.Id,
                    Sender = (message.Sender != null) ?
                              message.Sender.EmailAddress.Name :
                              "Unknown name",
                    From = (message.Sender != null) ?
                              message.Sender.EmailAddress.Address :
                              "Unknown email",
                    Subject = message.Subject ?? "No subject",
                    Importance = message.Importance.ToString()
                });
            }
            //var yyy = await _graphClient.Me.Drive.Root.Children.Request().GetAsync();
            //if (events.NextPageRequest != null)
            return MyMessages;
        }

        private static DateTime GetUtcStartOfWeekInTimeZone(DateTime today, TimeZoneInfo timeZone)
        {
            // Assumes Sunday as first day of week
            int diff = System.DayOfWeek.Sunday - today.DayOfWeek;

            // create date as unspecified kind
            var unspecifiedStart = DateTime.SpecifyKind(today.AddDays(diff), DateTimeKind.Unspecified);

            // convert to UTC
            return TimeZoneInfo.ConvertTimeToUtc(unspecifiedStart, timeZone);
        }
        // </GetCalendarViewSnippet>
    }
}
