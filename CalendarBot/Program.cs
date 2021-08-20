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
using System.IO;

using Excel = Microsoft.Office.Interop.Excel;

namespace CalendarBot
{
	class Program
	{
		static readonly DateTime today = DateTime.Today;

		readonly bool isWeekend = today.DayOfWeek == DayOfWeek.Sunday || today.DayOfWeek == DayOfWeek.Saturday;
		readonly string techFocus = "TFP";
		readonly string medFocus = "MFP";
		readonly string cinemaFocus = "CFP";
		readonly ulong tfpId = 747307858676416544, mfpId = 747307883653496862, cfpId = 747307901940531310;
		
		bool isTest = false;
		DiscordSocketClient client;

		static void Main()
		{
			new Program().MainAsync().GetAwaiter().GetResult();
		}
		
		async Task MainAsync() {
			await InitializeClient();

			if (!isWeekend)
			{
				SetSettings();
				await AnnounceEvents();
			}
			
			await AssignBirthdays();

			Console.ReadKey(true);
		}

		async Task InitializeClient()
		{
			DiscordSocketConfig config = new DiscordSocketConfig
			{
				AlwaysDownloadUsers = true
			};

			client = new DiscordSocketClient(config);

			client.Log += Log;

			string token = File.ReadAllText("D:/Documents/Calendar Bot Token.txt");

			await client.LoginAsync(TokenType.Bot, token);
			await client.StartAsync();

			//Wait for client to be connected
			while (client.ConnectionState != ConnectionState.Connected)
			{
				await Task.Delay(1000);
			}

			await Task.Delay(1000); //Allow for client to be ready after connected
		}

		void SetSettings()
		{
			Console.WriteLine("Is this a test? Y or N");
			char testAnswer = Console.ReadKey().KeyChar;
			isTest = char.ToUpper(testAnswer) == 'Y';
			Console.Write("\n");
		}

		async Task AnnounceEvents()
		{
			//download current master calendar
			WebClient webClient = new WebClient();
			byte[] calBytes = webClient.DownloadData("https://www.providencehigh.org/master/calendar/feed/ical.ics");
			string calString = Encoding.UTF8.GetString(calBytes);
			Calendar pHSCalendar = Calendar.Load(calString);

			//compile list of all of today's events
			List<CalendarEvent> todayEvents = new List<CalendarEvent>();
			foreach(CalendarEvent @event in pHSCalendar.Events)
			{
				if (EventIsToday(@event))
				{
					todayEvents.Add(@event);
				}
			}

			SocketGuild server = client.GetGuild(634944914756730880);

			//get correct channel for announcement
			ulong testChannelId = 812131884846415892;
			ulong calendarChannelId = 634945249864581150;
			ISocketMessageChannel channel = isTest ? server.GetTextChannel(testChannelId) : server.GetTextChannel(calendarChannelId);

			//get all roles for mention
			SocketRole calendarUpdatesRole = server.GetRole(740738668255379546);
			SocketRole tfpRole = server.GetRole(tfpId);
			SocketRole mfpRole = server.GetRole(mfpId);
			SocketRole cfpRole = server.GetRole(cfpId);

			//Form announcement message
			string announcementMessage = calendarUpdatesRole.Mention + "\nToday's events are:\n"; //intro
			Console.WriteLine(todayEvents.Count);
			if (todayEvents.Count == 0)
			{
				announcementMessage = calendarUpdatesRole.Mention + "\nThere are no events listed today.";
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

			//replace focus program acronyms with mentions for the focus program role
			announcementMessage = announcementMessage.Replace(techFocus, tfpRole.Mention);
			announcementMessage = announcementMessage.Replace(medFocus, mfpRole.Mention);
			announcementMessage = announcementMessage.Replace(cinemaFocus, cfpRole.Mention);

			Console.WriteLine(announcementMessage);

			await channel.SendMessageAsync(announcementMessage);
		}

		async Task AssignBirthdays()
		{
			Console.WriteLine("Assigning birthdays...");

			ulong bDayRoleId = 777806659556737024;

			SocketGuild server = client.GetGuild(634944914756730880);
			SocketRole bDayRole = server.GetRole(bDayRoleId);

			IReadOnlyCollection<IGuildUser> users = await (server as IGuild).GetUsersAsync();

			int j = 0;
			int memberCount = users.Count;

			foreach (IGuildUser user in users)
			{
				//If user has birthday role, remove it
				foreach (ulong roleId in user.RoleIds)
				{
					if (roleId == bDayRoleId)
					{
						await user.RemoveRoleAsync(bDayRole);
						continue;
					}
				}

				//print percent progrrss
				j++;
				int percentThrough = (j * 100) / memberCount;
				Console.WriteLine(user.Nickname + $": {percentThrough}%");
			}

			//Get correct cells for data
			var xlApp = new Excel.Application();
			var xlWorkbook = xlApp.Workbooks.Open(@"D:\OneDrive\OneDrive - Providence High School\Sophomore 20-21\Misc\Birthdays.xlsx");
			var xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];
			var xlRange = xlWorksheet.UsedRange;

			const int birthdayColumn = 2;
			int.TryParse((xlRange.Cells[2, 4] as Excel.Range).Value2.ToString(), out int rows); //Cell 2D has number of entries

			for (int i = 2; i <= rows + 1; i++)
			{
				/*
				Reading a Date from an Excel spreadsheet returns an integer of the number of days since December 29, 1889
				Not sure why, just go with it
				So add the day length to Decemeber 29, 1889 to get the true date
				*/
				DateTime day = new DateTime(1889, 12, 29);
				int.TryParse((xlRange.Cells[i, birthdayColumn] as Excel.Range).Value2.ToString(), out int dayLength);
				day = day.AddDays(dayLength);
				day = new DateTime(DateTime.Today.Year, day.Month, day.Day);

				if (day == today)
				{
					//Get user associated with that row
					string birthdayPerson = (xlRange.Cells[i, birthdayColumn - 1] as Excel.Range).Value2.ToString();
					IGuildUser birthdayUser = null;

					foreach (IGuildUser user in users)
					{
						//check each user to find person
						if (user.Username + "#" + user.Discriminator == birthdayPerson)
						{
							birthdayUser = user;
						}
					}

					await birthdayUser.AddRoleAsync(bDayRole);
				}
			}

			//close workbook and excel instance
			//only once instance may be open at a time
			xlWorkbook.Close(false);
			xlApp.Quit();

			Marshal.ReleaseComObject(xlWorksheet);
			Marshal.ReleaseComObject(xlWorkbook);
			Marshal.ReleaseComObject(xlApp);
		}

		bool EventIsToday(CalendarEvent eventInQuestion)
		{
			return (eventInQuestion.Start.ToTimeZone("America/Los_Angeles").Date <= today && today < eventInQuestion.End.ToTimeZone("America/Los_Angeles").Date) ||
					eventInQuestion.Start.ToTimeZone("America/Los_Angeles").Date == today.Date.ToLocalTime();
		}

		Task Log(LogMessage msg)
		{
			Console.WriteLine(msg.ToString());
			return Task.CompletedTask;
		}
	}
}