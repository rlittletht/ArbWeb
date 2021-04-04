using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using NUnit.Framework;
using OpenQA.Selenium.DevTools.V86.ServiceWorker;

namespace ArbWeb.Games
{
	public class LearnMappings
	{
		public class MapAndConfidences
		{
			private Dictionary<string, Dictionary<string, int>> mapAndConfidencesLeftToRight = new Dictionary<string, Dictionary<string, int>>();
			private Dictionary<string, Dictionary<string, int>> mapAndConfidencesRightToLeft = new Dictionary<string, Dictionary<string, int>>();

			/*----------------------------------------------------------------------------
				%%Function: Update
				%%Qualified: ArbWeb.Games.LearnMappings.MapAndConfidences.Update
			----------------------------------------------------------------------------*/
			static void Update(Dictionary<string, Dictionary<string, int>> mapAndConfidences, string sLeft, string sRight, int nConfidence)
			{
				if (!mapAndConfidences.ContainsKey(sLeft))
					mapAndConfidences.Add(sLeft, new Dictionary<string, int>());

				if (!mapAndConfidences[sLeft].ContainsKey(sRight))
					mapAndConfidences[sLeft].Add(sRight, 0);

				string[] rgsKeys = mapAndConfidences[sLeft].Keys.ToArray();
				
				foreach (string key in rgsKeys)
				{
					if (key == sRight)
						mapAndConfidences[sLeft][key] += nConfidence;
					else
						mapAndConfidences[sLeft][key] -= nConfidence;
				}
			}
			
			/*----------------------------------------------------------------------------
				%%Function: Update
				%%Qualified: ArbWeb.Games.LearnMappings.MapAndConfidences.Update
		
				When we get a left and right value, if we have never seen the values,
				add them.
		
				otherwise, if add 1 to the confidence of the two being associated,
				and subtract 1 for the confidence for all the other maps for them
			----------------------------------------------------------------------------*/
			public void Update(string sLeft, string sRight, int nConfidence)
			{
				sLeft = sLeft.ToUpper();
				sRight = sRight.ToUpper();

				Update(mapAndConfidencesLeftToRight, sLeft, sRight, nConfidence);
				Update(mapAndConfidencesRightToLeft, sRight, sLeft, nConfidence);
			}

			/*----------------------------------------------------------------------------
				%%Function: AddPotentialMap
				%%Qualified: ArbWeb.Games.LearnMappings.MapAndConfidences.AddPotentialMap

				We know we have an item that needs to be mapped, but we don't have any
				potential targets for it yet. Still record an empty target so we will
				be able to know if it never gets a target
			----------------------------------------------------------------------------*/
			static void AddPotentialMap(Dictionary<string, Dictionary<string, int>> mapAndConfidences, string sLeft)
			{
				if (!mapAndConfidences.ContainsKey(sLeft))
					mapAndConfidences.Add(sLeft, new Dictionary<string, int>());
			}

			/*----------------------------------------------------------------------------
				%%Function: AddPotentialMap
				%%Qualified: ArbWeb.Games.LearnMappings.MapAndConfidences.AddPotentialMap

				we have a value that needs to be mapped. we don't know if we have any
				candidates or not, but we know we have to map it...
			----------------------------------------------------------------------------*/
			public void AddPotentialMap(string sLeft)
			{
				sLeft = sLeft.ToUpper();
				AddPotentialMap(mapAndConfidencesLeftToRight, sLeft);
				AddPotentialMap(mapAndConfidencesRightToLeft, sLeft);
			}
			
			/*----------------------------------------------------------------------------
				%%Function: WriteConfidencesToFile
				%%Qualified: ArbWeb.Games.LearnMappings.MapAndConfidences.WriteConfidencesToFile
			----------------------------------------------------------------------------*/
			public static void WriteConfidencesToFile(string sOutFile, Dictionary<string, Dictionary<string, int>> mapAndConfidences)
			{
				using (StreamWriter sw = new StreamWriter(sOutFile, false, Encoding.Default))
				{
					sw.WriteLine($"Left,Right,Confidence");

					foreach (string keyLeft in mapAndConfidences.Keys)
					{
						foreach (string keyMatch in mapAndConfidences[keyLeft].Keys)
						{
							int confidence = mapAndConfidences[keyLeft][keyMatch];
							
							sw.WriteLine($"\"{keyLeft}\",\"{keyMatch}\",\"{confidence}\"");
						}
						
						if (mapAndConfidences[keyLeft].Keys.Count == 0)
							sw.WriteLine($"\"{keyLeft}\",\"***UNKNOWN***\",\"0\"");
					}
				}
			}

			public Dictionary<string, string> CreateMapFromLearning()
			{
				Dictionary<string, string> map = new Dictionary<string, string>();

				foreach (string keyLeft in mapAndConfidencesLeftToRight.Keys)
				{
					int nConfidenceHigh = 0;
					string sCurrentBest = null;

					foreach (string keyMatch in mapAndConfidencesLeftToRight[keyLeft].Keys)
					{
						int confidence = mapAndConfidencesLeftToRight[keyLeft][keyMatch];

						if (confidence > nConfidenceHigh)
						{
							sCurrentBest = keyMatch;
							nConfidenceHigh = confidence;
						}
					}

					if (sCurrentBest != null)
						map.Add(keyLeft, sCurrentBest);
				}

				return map;
			}

			/*----------------------------------------------------------------------------
				%%Function: WriteConfidencesToFile
				%%Qualified: ArbWeb.Games.LearnMappings.MapAndConfidences.WriteConfidencesToFile
			----------------------------------------------------------------------------*/
			public void WriteConfidencesToFile(string sOutfileLTR, string sOutfileRTL)
			{
				WriteConfidencesToFile(sOutfileLTR, mapAndConfidencesLeftToRight);
				WriteConfidencesToFile(sOutfileRTL, mapAndConfidencesRightToLeft);
			}
		}

		private MapAndConfidences siteMapAndConfidences = new MapAndConfidences();
		private MapAndConfidences teamMapAndConfidences = new MapAndConfidences();

		/*----------------------------------------------------------------------------
			%%Function: GroupGamesByDate
			%%Qualified: ArbWeb.Games.LearnMappings.GroupGamesByDate
		----------------------------------------------------------------------------*/
		static Dictionary<DateTime, List<SimpleGame>> GroupGamesByDate(SimpleSchedule schedule)
		{
			Dictionary<DateTime, List<SimpleGame>> gamesByDate = new Dictionary<DateTime, List<SimpleGame>>();
			
			foreach (SimpleGame game in schedule.Games)
			{
				if (!gamesByDate.ContainsKey(game.StartDateTime))
					gamesByDate.Add(game.StartDateTime, new List<SimpleGame>());

				gamesByDate[game.StartDateTime].Add(game);
			}

			return gamesByDate;
		}

		/*----------------------------------------------------------------------------
			%%Function: IsGameFuzzyMatch
			%%Qualified: ArbWeb.Games.LearnMappings.IsGameFuzzyMatch
		----------------------------------------------------------------------------*/
		static int IsGameFuzzyMatch(SimpleGame gameLeft, SimpleGame gameRight)
		{
			int nConfidenceFuzzySiteMatch = FuzzyMatcher.IsStringFuzzyMatch(gameLeft.Site, gameRight.Site);

			int numberLeft = Int32.Parse(gameLeft.Number);
			int numberRight = Int32.Parse(gameRight.Number);

			// if the numbers are an exact match, we only have to have low confidence
			// on the site matching to believe this is a certain match. otherwise, its just a 60%
			// match (all arbitrary numbers)
			if (numberLeft == numberRight)
				return nConfidenceFuzzySiteMatch >= 30 ? 100 : Math.Max(60, nConfidenceFuzzySiteMatch);
			
			int dNum = Math.Abs(numberLeft - numberRight);


			if (dNum == (dNum / 1000) * 1000)
				return nConfidenceFuzzySiteMatch >= 60 ? 100 : Math.Max(50, nConfidenceFuzzySiteMatch);

			return nConfidenceFuzzySiteMatch;
		}

		/*----------------------------------------------------------------------------
			%%Function: GenerateMapsFromSchedules
			%%Qualified: ArbWeb.Games.LearnMappings.GenerateMapsFromSchedules
		----------------------------------------------------------------------------*/
		public static ScheduleMaps GenerateMapsFromSchedules(SimpleSchedule scheduleLeft, SimpleSchedule scheduleRight)
		{
			// we have to find a list of games we think match... for now, that's just any game
			// with the same slot date/time.  hopefully that's enough to build a site map

			LearnMappings learner = new LearnMappings();
			
			// build a schedule grouped by DateTime
//			Dictionary<DateTime, List<SimpleGame>> gamesByTimeLeft = GroupGamesByDate(scheduleLeft);
			Dictionary<DateTime, List<SimpleGame>> gamesByTimeRight = GroupGamesByDate(scheduleRight);

			foreach (SimpleGame game in scheduleLeft.Games)
			{
				foreach (SimpleGame gameMatch in gamesByTimeRight[game.StartDateTime])
				{
					int nConfidence = IsGameFuzzyMatch(game, gameMatch);
					if (nConfidence > 0)
						learner.UpdateConfidencesForGames(game, gameMatch, nConfidence);

					// even if we weren't a match, we still have sites and teams that
					// need mapped (otherwise, we would never know about the teams/sites
					// that were only parts of games we couldn't match)
					learner.UpdatePossibleMaps(game);
				}
			}

			#if false

			learner.WriteLearningsToFile(
				@"c:\temp\sitemapsLTR.csv",
				@"c:\temp\sitemapsRTL.csv",
				@"c:\temp\teammapsLTR.csv",
				@"c:\temp\teammapsRTL.csv");
			#endif

			return new ScheduleMaps(learner.teamMapAndConfidences, learner.siteMapAndConfidences);
		}

		/*----------------------------------------------------------------------------
			%%Function: WriteLearningsToFile
			%%Qualified: ArbWeb.Games.LearnMappings.WriteLearningsToFile
		----------------------------------------------------------------------------*/
		public void WriteLearningsToFile(string sSiteFileLTR, string sSiteFileRTL, string sTeamFileLTR, string sTeamFileRTL)
		{
			siteMapAndConfidences.WriteConfidencesToFile(sSiteFileLTR, sSiteFileRTL);
			teamMapAndConfidences.WriteConfidencesToFile(sTeamFileLTR, sTeamFileRTL);
		}

		/*----------------------------------------------------------------------------
			%%Function: UpdateConfidencesForGames
			%%Qualified: ArbWeb.Games.LearnMappings.UpdateConfidencesForGames
		----------------------------------------------------------------------------*/
		public void UpdateConfidencesForGames(SimpleGame gameLeft, SimpleGame gameRight, int nConfidence)
		{
			siteMapAndConfidences.Update(gameLeft.Site, gameRight.Site, nConfidence);
			teamMapAndConfidences.Update(gameLeft.Home, gameRight.Home, nConfidence);
			teamMapAndConfidences.Update(gameLeft.Away, gameRight.Away, nConfidence);
		}

		/*----------------------------------------------------------------------------
			%%Function: UpdatePossibleMaps
			%%Qualified: ArbWeb.Games.LearnMappings.UpdatePossibleMaps
		----------------------------------------------------------------------------*/
		public void UpdatePossibleMaps(SimpleGame game)
		{
			siteMapAndConfidences.AddPotentialMap(game.Site);
			teamMapAndConfidences.AddPotentialMap(game.Home);
			teamMapAndConfidences.AddPotentialMap(game.Away);
		}
	}
}