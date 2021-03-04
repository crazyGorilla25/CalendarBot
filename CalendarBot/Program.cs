using Discord;
using Discord.WebSocket;
using Ical.Net;
using Ical.Net.CalendarComponents;
using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace CalendarBot
{
	class Program
    {
        readonly string techFocus = "TFP";
        readonly string medFocus = "MFP";
        readonly string cinemaFocus = "CFP";

        readonly ulong tfpId = 747307858676416544, mfpId = 747307883653496862, cfpId = 747307901940531310;
        DiscordSocketClient client;
        static bool normalTime = false;
        
        static void Main()
        {
            Console.WriteLine("Run at normal time? Y or N");
            char answer = Console.ReadKey().KeyChar;
            normalTime = answer == 'y' || answer == 'Y';
            Console.Write("\n");
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
            DateTime timeToRun;
            if (normalTime)
            {
                timeToRun = new DateTime(now.Year, now.Month, now.Day, 8, 0, 0);
            }
            else
            {
                timeToRun = now.AddSeconds(5);
            }
            TimeSpan delay;
			if (timeToRun <= now)
			{
                timeToRun = timeToRun.AddDays(1);
			}
            delay = timeToRun - now;

            Console.WriteLine(delay.Hours + " hours, " + delay.Minutes + " minutes, & " + delay.Seconds + " seconds");

            WebClient webClient = new WebClient();
            byte[] calBytes = webClient.DownloadData("https://www.providencehigh.org/master/calendar/feed/ical.ics");
            string calString = Encoding.UTF8.GetString(calBytes);
            Calendar pHSCalendar = Calendar.Load(calString);
            List<CalendarEvent> todayEvents = new List<CalendarEvent>();

            Console.WriteLine(DateTime.Today);

            for (int i = 0; i < pHSCalendar.Events.Count; i++)
            {

                CalendarEvent eventInQuestion = pHSCalendar.Events[i];
                //Console.WriteLine(eventInQuestion.Summary + " Start: " + eventInQuestion.Start.ToTimeZone("America/Los_Angeles").Date + " End: " + eventInQuestion.End.ToTimeZone("America/Los_Angeles").Date);
                if (EventIsToday(eventInQuestion))
                {
                    todayEvents.Add(eventInQuestion);
                }
            }

            SocketGuild server = client.GetGuild(634944914756730880);
            ISocketMessageChannel channel = server.GetTextChannel(634945249864581150);

            SocketRole role = server.GetRole(740738668255379546);
            SocketRole tfpRole = server.GetRole(tfpId);
            SocketRole mfpRole = server.GetRole(mfpId);
            SocketRole cfpRole = server.GetRole(cfpId);

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
                    string summary = todayEvents[i].Summary;
                    summary = summary.Replace(techFocus, tfpRole.Mention);
                    summary = summary.Replace(medFocus, mfpRole.Mention);
                    summary = summary.Replace(cinemaFocus, cfpRole.Mention);

                    if (i != todayEvents.Count - 1)
                    {
                        announcementMessage += "**" + summary + "**, \n\n";
                    }
                    else
                    {
                        announcementMessage += "and **" + summary + "**";
                    }
                }
            }

            announcementMessage += "\n\nSubmit your attendance here: https://bit.ly/3b9YQ0L";

            Console.WriteLine(announcementMessage);


            await Task.Delay(delay);

            await channel.SendMessageAsync(announcementMessage);
        }

        bool EventIsToday(CalendarEvent eventInQuestion)
		{
            return (eventInQuestion.Start.ToTimeZone("America/Los_Angeles").Date <= DateTime.Today && DateTime.Today < eventInQuestion.End.ToTimeZone("America/Los_Angeles").Date) ||
                    eventInQuestion.Start.ToTimeZone("America/Los_Angeles").Date == DateTime.Today.Date.ToLocalTime();

        }
    }
}