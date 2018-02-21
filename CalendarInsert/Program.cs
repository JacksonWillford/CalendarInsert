using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Globalization;

namespace CalendarInsert
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/calendar-dotnet-quickstart.json
        static string[] Scopes = { CalendarService.Scope.Calendar };
        static string ApplicationName = "Google Calendar API .NET Quickstart";

        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/calendar-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            string date = "";
            string time = "";

            string[] timeString;
            List<string> job = new List<string>();
            List<string> startTimeFormatted = new List<string>();
            List<string> endTimeFormatted = new List<string>();
            List<int> startingHour = new List<int>();
            DateTime startDate;
            DateTime endDate;
            try
            {
                string[] scheduleRaw = File.ReadAllLines(@"schedule.txt");
                for (int i = 0; i < scheduleRaw.Count(); i++)
                {
                    if (i % 3 == 0)
                    {
                        date = scheduleRaw[i];
                        job.Add(scheduleRaw[i + 1]);
                        time = scheduleRaw[i + 2];
                        timeString = time.Split('-');
                        startDate = DateTime.ParseExact(date + " " + timeString[0], "MM/dd/yy h:mmtt ", CultureInfo.InvariantCulture);
                        endDate = DateTime.ParseExact(date + timeString[1], "MM/dd/yy h:mmtt", CultureInfo.InvariantCulture);

                        if (startDate.ToString("HH") == "00" || startDate.ToString("HH") == "01" || startDate.ToString("HH") == "02" || startDate.ToString("HH") == "03")
                        {
                            startDate = startDate.AddDays(1);
                        }
                        startTimeFormatted.Add(startDate.ToString("yyyy-MM-dd" + "T" + "HH':'mm':'ss"));
                        if (endDate.ToString("HH") == "00" || endDate.ToString("HH") == "01" || endDate.ToString("HH") == "02" || endDate.ToString("HH") == "03")
                        {
                            endDate = endDate.AddDays(1);
                        }
                        endTimeFormatted.Add(endDate.ToString("yyyy-MM-dd" + "T" + "HH':'mm':'ss"));
                        startingHour.Add(Convert.ToInt32(startDate.Hour));
                    }
                }
            }
            catch (FormatException)
            {
                Console.WriteLine("File was not found!");
            }
            catch (IndexOutOfRangeException)
            {
                Console.WriteLine("Format Exception!");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            //for (int i = 0; i < job.Count(); i++)
            //{
            //    Console.WriteLine("Start: " + startTimeFormatted[i]);

            //    Console.WriteLine("End:   " + endTimeFormatted[i]);
            //}
            //for (int i = 0; i < job.Count(); i++)
            //{
            //    Console.WriteLine(job[i]);
            //    Console.WriteLine(startTimeFormatted[i]);
            //    Console.WriteLine(endTimeFormatted[0]);
            //}
            for (int i = 0; i < job.Count(); i++)
            {
                Event workDay = new Event()
                {
                    Summary = job[i],
                    Location = "4211 Trueman Blvd. Hilliard, OH 43026",
                    Start = new EventDateTime()
                    {
                        DateTime = DateTime.Parse(startTimeFormatted[i]),
                        TimeZone = "America/New_York"
                    },
                    End = new EventDateTime()
                    {
                        DateTime = DateTime.Parse(endTimeFormatted[i]),
                        TimeZone = "America/New_York"
                    },
                };
                if (startingHour[i] <= 12)
                {
                    workDay.Reminders = new Event.RemindersData()
                    {
                        UseDefault = false,
                        Overrides = new EventReminder[]
                        {
                            new EventReminder() {Method = "popup", Minutes = 480 },
                            new EventReminder() {Method = "popup", Minutes = 120 },
                            new EventReminder() {Method = "popup", Minutes = 60 },
                        },
                    };
                }
                else
                {
                    workDay.Reminders = new Event.RemindersData()
                    {
                        UseDefault = false,
                        Overrides = new EventReminder[]
                        {
                            new EventReminder() {Method = "popup", Minutes = 120 },
                            new EventReminder() {Method = "popup", Minutes = 60 },
                        },
                    };
                }
                string calendarId = "mdhgaj2lfkkob7iqdvo17b6nl0@group.calendar.google.com";
                EventsResource.InsertRequest request = service.Events.Insert(workDay, calendarId);
                try
                {
                    Event createdEvent = request.Execute();
                    Console.WriteLine("Event created: {0}", createdEvent.HtmlLink);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
            Console.WriteLine("This window will close in 5 seconds");
            Thread.Sleep(5000);
        }
    }
}