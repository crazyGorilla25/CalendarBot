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
        ISocketMessageChannel channelToAnnounce;
        
        static void Main(string[] args)
        {
            new Program().MainAsync().GetAwaiter().GetResult();
        }
        
        public async Task MainAsync() {
            client = new DiscordSocketClient();

            client.Log += Log;

            //  You can assign your bot token to a string, and pass that in to connect.
            //  This is, however, insecure, particularly if you plan to have your code hosted in a public repository.
            var token = "NzQwNjYxNjAwMTg4NDk4MTAx.XysQ3g.GX7_U6pYzDq3u1UVkTCqv5LXtq";

            // Some alternative options would be to keep your token in an Environment Variable or a standalone file.
            // var token = Environment.GetEnvironmentVariable("NameOfYourEnvironmentVariable");
            // var token = File.ReadAllText("token.txt");
            // var token = JsonConvert.DeserializeObject<AConfigurationClass>(File.ReadAllText("config.json")).Token;

            await client.LoginAsync(TokenType.Bot, token);
            await client.StartAsync();

            client.MessageReceived += MessageRecieved;
            await AnnounceEvents();

            // Block this task until the program is closed.
            await Task.Delay(-1);
        }

        private async Task MessageRecieved(SocketMessage message) {
            if(message.Content == "*setChannelAnnounce")
            {
                channelToAnnounce = message.Channel;
                await message.Channel.SendMessageAsync("Announce channel set.");
            }
        }

        private Task Log(LogMessage msg)
        {
            Console.WriteLine(msg.ToString());
            return Task.CompletedTask;
        }

        async Task AnnounceEvents()
        {
            DateTime now = DateTime.Now;
            DateTime timeToRun = new DateTime(now.Year, now.Month, now.Day, 7, 0, 0);
            TimeSpan delay;
            if(timeToRun > now)
            {
                delay = timeToRun - now;
            }
            else
            {
                timeToRun.AddDays(1);
                delay = timeToRun - now;
            }

            await Task.Delay(delay);

            WebClient webClient = new WebClient();
            byte[] calBytes = webClient.DownloadData("https://www.providencehigh.org/master/calendar/feed/ical.ics");
            string calString = Encoding.UTF8.GetString(calBytes);
            Calendar pHSCalendar = Calendar.Load(calString);
            List<CalendarEvent> todayEvents = new List<CalendarEvent>();
            for (int i = 0; i < pHSCalendar.Events.Count; i++)
            {
                CalendarEvent eventInQuestion = pHSCalendar.Events[i];
                if (eventInQuestion.Start.Date == DateTime.Today)
                {
                    todayEvents.Add(eventInQuestion);
                }
            }
            if (channelToAnnounce != null)
            {
                await channelToAnnounce.SendMessageAsync(todayEvents[0].Summary);
            }
        }
    }
}