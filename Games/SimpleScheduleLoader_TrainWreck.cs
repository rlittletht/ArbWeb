using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using TCore.CSV;

namespace ArbWeb.Games
{
	public class SimpleScheduleLoader_TrainWreck
	{
		/*----------------------------------------------------------------------------
			%%Function: LoadFromExcelFile
			%%Qualified: ArbWeb.Games.SimpleScheduleLoader_TrainWreck.LoadFromExcelFile
		----------------------------------------------------------------------------*/
		public static SimpleSchedule LoadFromExcelFile(string sExcelFile)
		{
			string sCsvFile = $"{Environment.GetEnvironmentVariable("Temp")}\\temp7140c22b-029e-45d0-b3da-8a3061c13d8b.csv"; //temp{System.Guid.NewGuid().ToString()}.csv";

//			DownloadGenericExcelReport.ConvertExcelFileToCsv(sExcelFile, sCsvFile);

			return LoadFromCsvFile(sCsvFile);
		}

		class ColumnInfo
		{
			public int StartDateTime { get; set; }
			public int StartDate { get; set; }
			public int StartTime { get; set; }
			public int Site { get; set; }
			public int Level { get; set; }
			public int Home { get; set; }
			public int Away { get; set; }
			public int Number { get; set; }

			int LookupColumnForHeaderStrings(CsvFile csvFile, IEnumerable<string> headingNames)
			{
				foreach (string headingName in headingNames)
				{
					int i = csvFile.LookupColumn(headingName);
					if (i != -1)
						return i;
				}

				return -1;
			}
			
			public ColumnInfo(CsvFile csvFile)
			{
				StartTime = LookupColumnForHeaderStrings(csvFile, new[] {"Time", "StartTime", "Start Time", "GameTime", "Game Time"});
				if (StartTime != -1)
				{
					// we have a start time, that means the date column is just the date
					StartDate = LookupColumnForHeaderStrings(csvFile, new[] { "Date", "StartDate", "Start Date", "GameDate", "Game Date" });;
					if (StartDate == -1)
						throw new Exception($"Time column without date column");
					StartDateTime = -1;
				}
				else
				{
					StartDateTime = LookupColumnForHeaderStrings(csvFile, new[] { "Date", "StartDate", "Start Date", "GameDate", "Game Date", "DateTime", "Date Time", "GameDateTime", "Game DateTime", "Game Date Time" });
					if (StartDateTime == -1)
						throw new Exception($"No date/time column for game");
					StartDate = -1;
					StartTime = -1;
				}
				
				Site = LookupColumnForHeaderStrings(csvFile, new[] { "Site", "Field", "Location" });
				Level = LookupColumnForHeaderStrings(csvFile, new[] { "Level", "Division" });
				Home = LookupColumnForHeaderStrings(csvFile, new[] { "Home", "Home Team" });
				Away = LookupColumnForHeaderStrings(csvFile, new[] { "Away", "Away Team", "Visitor" });
				Number = LookupColumnForHeaderStrings(csvFile, new[] { "Game", "GameTag", "Number", "Game Number", "GameID" });

				if (Site == -1 || Level == -1 || Home == -1 || Away == -1 || Number == -1)
					throw new Exception("couldn't find required column");
			}
		}
		
		/*----------------------------------------------------------------------------
			%%Function: LoadFromCsvFile
			%%Qualified: ArbWeb.Games.SimpleScheduleLoader_TrainWreck.LoadFromCsvFile
		----------------------------------------------------------------------------*/
		public static SimpleSchedule LoadFromCsvFile(string sCsvFile)
		{
			CsvFile csvFile = new CsvFile(sCsvFile);
			SimpleSchedule schedule = new SimpleSchedule();
			
			csvFile.ReadHeadingLine();
			ColumnInfo columnInfo = new ColumnInfo(csvFile);
			
			// let's make sure we have all the fields we need to have (be tolerant for heading aliases)

			while (csvFile.ReadNextCsvLine())
			{
				DateTime startDateTime;

				if (columnInfo.StartDateTime != -1)
				{
					startDateTime = DateTime.Parse(csvFile.GetValue(columnInfo.StartDateTime));
				}
				else
				{
					DateTime startDate = DateTime.Parse(csvFile.GetValue(columnInfo.StartDate));
					DateTime startTime = DateTime.Parse(csvFile.GetValue(columnInfo.StartTime));

					startDateTime = startDate.Add(startTime.TimeOfDay);
				}

				SimpleGame game = new SimpleGame(
					startDateTime,
					csvFile.GetValue(columnInfo.Site),
					csvFile.GetValue(columnInfo.Level),
					csvFile.GetValue(columnInfo.Home),
					csvFile.GetValue(columnInfo.Away),
					csvFile.GetValue(columnInfo.Number));

				schedule.AddSimpleGame(game);
			}

			return schedule;
		}
	}
}