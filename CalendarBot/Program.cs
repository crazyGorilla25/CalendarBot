using Discord;
using Discord.WebSocket;
using Ical.Net;
using Ical.Net.CalendarComponents;
using System;
using System.Collections.Generic;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CalendarBot
{
	class Program
	{
		static readonly DateTime today = DateTime.Today;
		static readonly bool isWeekend = !(today.DayOfWeek != DayOfWeek.Sunday && today.DayOfWeek != DayOfWeek.Saturday);

		static bool runAtNormalTime = false, isTest = false;

		readonly string techFocus = "TFP";
		readonly string medFocus = "MFP";
		readonly string cinemaFocus = "CFP";
		readonly ulong tfpId = 747307858676416544, mfpId = 747307883653496862, cfpId = 747307901940531310;

		DiscordSocketClient client;

		static void Main()
		{
			if (!isWeekend && today >= new DateTime(2021, 8, 19))
			{
				Console.WriteLine("Is this a test? Y or N");
				char testAnswer = Console.ReadKey().KeyChar;
				isTest = char.ToUpper(testAnswer) == 'Y';
				Console.Write("\n");

				Console.WriteLine("Run at normal time? Y or N");
				char answer = Console.ReadKey().KeyChar;
				runAtNormalTime = char.ToUpper(answer) == 'Y';
				Console.Write("\n");

			}
			new Program().MainAsync().GetAwaiter().GetResult();
		}
		
		async Task MainAsync() {
			DiscordSocketConfig config = new DiscordSocketConfig
			{
				AlwaysDownloadUsers = true
			};

			client = new DiscordSocketClient(config);

			client.Log += Log;

			string token = "NzQwNjYxNjAwMTg4NDk4MTAx.XysQ3g.5mk4KZEzX5dESJEjqFAsGwVLvd8";

			await client.LoginAsync(TokenType.Bot, token);
			await client.StartAsync();

			if(!isWeekend && today >= new DateTime(2021, 8, 19)) await AnnounceEvents();

			Console.WriteLine("Assigning birthdays...");

			await Task.Delay(new TimeSpan(0, 0, 5));
			
			await AssignBirthdays();
		}

		Task Log(LogMessage msg)
		{
			Console.WriteLine(msg.ToString());
			return Task.CompletedTask;
		}

		async Task AssignBirthdays()
		{
			SocketGuild server = client.GetGuild(634944914756730880);
			SocketRole bDayRole = server.GetRole(777806659556737024);

			IReadOnlyCollection<IGuildUser> users = await (server as IGuild).GetUsersAsync();

			int j = 0;
			int memberCount = users.Count;

			foreach(IGuildUser user in users)
			{
				await user.RemoveRoleAsync(bDayRole);
				j++;
				int percentThrough = (j*100)/memberCount;
				Console.WriteLine(user.Nickname + $": {percentThrough}%");
			}

			var xlApp = new Excel.Application();
			var xlWorkbook = xlApp.Workbooks.Open(@"D:\OneDrive\OneDrive - Providence High School\Sophomore 20-21\Misc\Birthdays.xlsx");
			var xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];
			var xlRange = xlWorksheet.UsedRange;

			const int birthdayColumn = 2;
			int.TryParse((xlRange.Cells[2, 4] as Excel.Range).Value2.ToString(), out int rows);

			for (int i = 2; i <= rows+1; i++)
			{
				DateTime day = new DateTime(1889, 12, 29);
				int.TryParse((xlRange.Cells[i, birthdayColumn] as Excel.Range).Value2.ToString(), out int dayLength);
				day = day.AddDays(dayLength);
				day = new DateTime(DateTime.Today.Year, day.Month, day.Day);

				if(day == today)
				{
					string birthdayPerson = (xlRange.Cells[i, birthdayColumn - 1] as Excel.Range).Value2.ToString();
					IGuildUser birthdayUser = null;

					foreach (IGuildUser user in users)
					{
						if(user.Username + "#" + user.Discriminator == birthdayPerson)
						{
							birthdayUser = user;
						}
					}

					await birthdayUser.AddRoleAsync(bDayRole);
				}
			}

			xlWorkbook.Close(false);
			xlApp.Quit();

			Marshal.ReleaseComObject(xlWorksheet);
			Marshal.ReleaseComObject(xlWorkbook);
			Marshal.ReleaseComObject(xlApp);
		}

		async Task AnnounceEvents()
		{
			Console.WriteLine("Announcing events");
			DateTime now = DateTime.Now;
			DateTime timeToRun;
			if (runAtNormalTime)
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
				if (EventIsToday(eventInQuestion))
				{
					todayEvents.Add(eventInQuestion);
				}
			}


			SocketGuild server = client.GetGuild(634944914756730880);
			ISocketMessageChannel channel = isTest ? server.GetTextChannel(812131884846415892) : server.GetTextChannel(634945249864581150);

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
			return (eventInQuestion.Start.ToTimeZone("America/Los_Angeles").Date <= today && today < eventInQuestion.End.ToTimeZone("America/Los_Angeles").Date) ||
					eventInQuestion.Start.ToTimeZone("America/Los_Angeles").Date == today.Date.ToLocalTime();
		}
	}
}