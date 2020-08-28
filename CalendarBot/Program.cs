using System;
using System.Text;
using System.Net;
using System.Collections.Generic;
using System.Threading.Tasks;

using Ical.Net;
using Ical.Net.CalendarComponents;

using Discord;
using Discord.WebSocket;

namespace CalendarBot
{
    class Program
    {
        private DiscordSocketClient client;
        
        static void Main()
        {
            new Program().MainAsync().GetAwaiter().GetResult();
        }
        
        async Task MainAsync() {
            client = new DiscordSocketClient();

            client.Log += Log;

            var token = "NzQwNjYxNjAwMTg4NDk4MTAx.XysQ3g.5mk4KZEzX5dESJEjqFAsGwVLvd8";

            await client.LoginAsync(TokenType.Bot, token);
            await client.StartAsync();
            await AnnounceEvents();
        }

        Task Log(LogMessage msg)
        {
            Console.WriteLine(msg.ToString());
            return Task.CompletedTask;
        }

        async Task AnnounceEvents()
        {
            Console.WriteLine("Announcing events");
            DateTime now = DateTime.Now;
            DateTime timeToRun = new DateTime(now.Year, now.Month, now.Day, 8, 0, 0);
            //DateTime timeToRun = now.AddSeconds(5);
            TimeSpan delay;
            if(timeToRun > now)
            {
                delay = timeToRun - now;
            }
            else
            {
                timeToRun = timeToRun.AddDays(1);
                delay = timeToRun - now;
            }

            await Task.Delay(delay);

            WebClient webClient = new WebClient();
            byte[] calBytes = webClient.DownloadData("https://www.providencehigh.org/master/calendar/feed/ical.ics");
            string calString = Encoding.UTF8.GetString(calBytes);
            Calendar pHSCalendar = Calendar.Load(calString);
            List<CalendarEvent> todayEvents = new List<CalendarEvent>();
            Console.WriteLine(DateTime.Today);
            for (int i = 0; i < pHSCalendar.Events.Count; i++)
            {

                CalendarEvent eventInQuestion = pHSCalendar.Events[i];
                Console.WriteLine(eventInQuestion.Summary + " Start: " + eventInQuestion.Start.Date + " End: " + eventInQuestion.End.Date);
                if ((eventInQuestion.Start.Date <= DateTime.Today && DateTime.Today < eventInQuestion.End.Date) || eventInQuestion.Start.Date == DateTime.Today.Date)
                {
                    todayEvents.Add(eventInQuestion);
                }
            }

            SocketGuild server = client.GetGuild(634944914756730880);
            ISocketMessageChannel channel = server.GetTextChannel(634945249864581150);

            SocketRole role = server.GetRole(740738668255379546);
            string announcementMessage = role.Mention + "\nToday's events are:\n";
            Console.WriteLine(todayEvents.Count);
            if (todayEvents.Count == 0)
            {
                announcementMessage = role.Mention + "\nThere are no events listed today.";
            }
            else if (todayEvents.Count == 1)
            {
                announcementMessage += "**" + todayEvents[0].Summary + "**";
            }
            else
            {
                for (int i = 0; i < todayEvents.Count; i++)
                {
                    if (i != todayEvents.Count - 1)
                    {
                        announcementMessage += "**" + todayEvents[i].Summary + "**, \n\n";
                    }
                    else
                    {
                        announcementMessage += "and **" + todayEvents[i].Summary + "**";
                    }
                }
            }

            await channel.SendMessageAsync(announcementMessage);
        }
    }
}