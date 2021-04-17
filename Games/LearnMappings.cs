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
			public class ConfidenceVal
			{
				public int Opportunities { get; set; }
				public int Cumulative { get; set; }
				public int Confidence => (Cumulative > 0 ? Cumulative / Opportunities : 0);
			}

			// pretty maps allow us to take the canonicalized name and map it back to
			// the original friendly version (mixed case, likely).
			private Dictionary<string, string> m_mapLeftPrettyName = new Dictionary<string, string>();
			private Dictionary<string, string> m_mapRightPrettyName = new Dictionary<string, string>();

			private Dictionary<string, Dictionary<string, ConfidenceVal>> m_mapAndConfidencesLeftToRight = new Dictionary<string, Dictionary<string, ConfidenceVal>>();
			private Dictionary<string, Dictionary<string, ConfidenceVal>> m_mapAndConfidencesRightToLeft = new Dictionary<string, Dictionary<string, ConfidenceVal>>();

			/*----------------------------------------------------------------------------
				%%Function: UpdatePrettyMap
				%%Qualified: ArbWeb.Games.LearnMappings.MapAndConfidences.UpdatePrettyMap
			----------------------------------------------------------------------------*/
			static void UpdatePrettyMap(Dictionary<string, string> mapPretty, string sPrettyName)
			{
				string sUpper = sPrettyName.ToUpper();

				if (mapPretty.ContainsKey(sUpper))
					mapPretty[sUpper] = sPrettyName;
				else
					mapPretty.Add(sUpper, sPrettyName);
			}
			
			public void UpdateLeftPretty(string sPretty) => UpdatePrettyMap(m_mapLeftPrettyName, sPretty);
			public void UpdateRightPretty(string sPretty) => UpdatePrettyMap(m_mapRightPrettyName, sPretty);

			public string LeftPretty(string sUgly) => m_mapLeftPrettyName[sUgly];
			public string RightPretty(string sUgly) => m_mapRightPrettyName[sUgly];

			/*----------------------------------------------------------------------------
				%%Function: Update
				%%Qualified: ArbWeb.Games.LearnMappings.MapAndConfidences.Update
			----------------------------------------------------------------------------*/
			static void Update(Dictionary<string, Dictionary<string, ConfidenceVal>> mapAndConfidences, string sLeft, string sRight, int nConfidence)
			{
				if (!mapAndConfidences.ContainsKey(sLeft))
					mapAndConfidences.Add(sLeft, new Dictionary<string, ConfidenceVal>());

				if (!mapAndConfidences[sLeft].ContainsKey(sRight))
					mapAndConfidences[sLeft].Add(sRight, new ConfidenceVal());

				string[] rgsKeys = mapAndConfidences[sLeft].Keys.ToArray();
				
				foreach (string key in rgsKeys)
				{
					if (key == sRight)
					{
						mapAndConfidences[sLeft][key].Cumulative += nConfidence;
						mapAndConfidences[sLeft][key].Opportunities++;
					}
					else
					{
						mapAndConfidences[sLeft][key].Cumulative -= nConfidence;
						mapAndConfidences[sLeft][key].Opportunities++;
					}
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
				UpdateLeftPretty(sLeft);
				UpdateRightPretty(sRight);
				
				sLeft = sLeft.ToUpper();
				sRight = sRight.ToUpper();

				Update(m_mapAndConfidencesLeftToRight, sLeft, sRight, nConfidence);
				Update(m_mapAndConfidencesRightToLeft, sRight, sLeft, nConfidence);
			}

			/*----------------------------------------------------------------------------
				%%Function: AddPotentialMap
				%%Qualified: ArbWeb.Games.LearnMappings.MapAndConfidences.AddPotentialMap

				We know we have an item that needs to be mapped, but we don't have any
				potential targets for it yet. Still record an empty target so we will
				be able to know if it never gets a target
			----------------------------------------------------------------------------*/
			static void AddPotentialMap(Dictionary<string, Dictionary<string, ConfidenceVal>> mapAndConfidences, string sLeft)
			{
				if (!mapAndConfidences.ContainsKey(sLeft))
					mapAndConfidences.Add(sLeft, new Dictionary<string, ConfidenceVal>());
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
				AddPotentialMap(m_mapAndConfidencesLeftToRight, sLeft);
				AddPotentialMap(m_mapAndConfidencesRightToLeft, sLeft);
			}
			
			/*----------------------------------------------------------------------------
				%%Function: WriteConfidencesToFile
				%%Qualified: ArbWeb.Games.LearnMappings.MapAndConfidences.WriteConfidencesToFile
			----------------------------------------------------------------------------*/
			public static void WriteConfidencesToFile(string sOutFile, Dictionary<string, Dictionary<string, ConfidenceVal>> mapAndConfidences)
			{
				using (StreamWriter sw = new StreamWriter(sOutFile, false, Encoding.Default))
				{
					sw.WriteLine($"Left,Right,Confidence");

					foreach (string keyLeft in mapAndConfidences.Keys)
					{
						foreach (string keyMatch in mapAndConfidences[keyLeft].Keys)
						{
							ConfidenceVal confidence = mapAndConfidences[keyLeft][keyMatch];
							
							sw.WriteLine($"\"{keyLeft}\",\"{keyMatch}\",\"{confidence.Cumulative}\",\"{confidence.Confidence}");
						}
						
						if (mapAndConfidences[keyLeft].Keys.Count == 0)
							sw.WriteLine($"\"{keyLeft}\",\"***UNKNOWN***\",\"0\",\"0\"");
					}
				}
			}

			/*----------------------------------------------------------------------------
				%%Function: CreateMapFromLearning
				%%Qualified: ArbWeb.Games.LearnMappings.MapAndConfidences.CreateMapFromLearning
			----------------------------------------------------------------------------*/
			public Dictionary<string, string> CreateMapFromLearning()
			{
				Dictionary<string, string> map = new Dictionary<string, string>();

				foreach (string keyLeft in m_mapAndConfidencesLeftToRight.Keys)
				{
					int nConfidenceHigh = 0;
					string sCurrentBest = null;

					foreach (string keyMatch in m_mapAndConfidencesLeftToRight[keyLeft].Keys)
					{
						int confidence = m_mapAndConfidencesLeftToRight[keyLeft][keyMatch].Confidence;
						
						if (confidence > nConfidenceHigh)
						{
							sCurrentBest = keyMatch;
							nConfidenceHigh = confidence;
						}
					}

					if (sCurrentBest != null)
						map.Add(keyLeft, RightPretty(sCurrentBest));
				}

				return map;
			}

			/*----------------------------------------------------------------------------
				%%Function: CreateMapFromLearning
				%%Qualified: ArbWeb.Games.LearnMappings.MapAndConfidences.CreateMapFromLearning
			----------------------------------------------------------------------------*/
			public Dictionary<string, string> CreateReverseMapFromLearning()
			{
				Dictionary<string, string> map = new Dictionary<string, string>();

				foreach (string keyRight in m_mapAndConfidencesRightToLeft.Keys)
				{
					int nConfidenceHigh = 0;
					string sCurrentBest = null;

					foreach (string keyMatch in m_mapAndConfidencesRightToLeft[keyRight].Keys)
					{
						int confidence = m_mapAndConfidencesRightToLeft[keyRight][keyMatch].Confidence;

						if (confidence > nConfidenceHigh)
						{
							sCurrentBest = keyMatch;
							nConfidenceHigh = confidence;
						}
					}

					if (sCurrentBest != null)
						map.Add(keyRight, LeftPretty(sCurrentBest));
				}

				return map;
			}

			/*----------------------------------------------------------------------------
				%%Function: WriteConfidencesToFile
				%%Qualified: ArbWeb.Games.LearnMappings.MapAndConfidences.WriteConfidencesToFile
			----------------------------------------------------------------------------*/
			public void WriteConfidencesToFile(string sOutfileLTR, string sOutfileRTL)
			{
				WriteConfidencesToFile(sOutfileLTR, m_mapAndConfidencesLeftToRight);
				WriteConfidencesToFile(sOutfileRTL, m_mapAndConfidencesRightToLeft);
			}
		}

		private MapAndConfidences m_siteMapAndConfidences = new MapAndConfidences();
		private MapAndConfidences m_teamMapAndConfidences = new MapAndConfidences();
		private MapAndConfidences m_gameTagMapAndConfidences = new MapAndConfidences();
		
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
				// even if we never match, we still have sites and teams that
				// need mapped (otherwise, we would never know about the teams/sites
				// that were only parts of games we couldn't match)
				learner.UpdatePossibleMaps(game);
				
				if (!gamesByTimeRight.ContainsKey(game.StartDateTime))
					continue;
				
				foreach (SimpleGame gameMatch in gamesByTimeRight[game.StartDateTime])
				{
					int nConfidence = FuzzyMatcher.IsGameFuzzyMatch(game, gameMatch);
					if (nConfidence > 0)
						learner.UpdateConfidencesForGames(game, gameMatch, nConfidence);
				}
			}

			learner.WriteLearningsToFile(
				@"c:\temp\sitemapsLTR.csv",
				@"c:\temp\sitemapsRTL.csv",
				@"c:\temp\teammapsLTR.csv",
				@"c:\temp\teammapsRTL.csv",
				@"c:\temp\numbermapLTR.csv",
				@"c:\temp\numbermapRTL.csv");

			return new ScheduleMaps(learner.m_teamMapAndConfidences, learner.m_siteMapAndConfidences, learner.m_gameTagMapAndConfidences);
		}

		/*----------------------------------------------------------------------------
			%%Function: WriteLearningsToFile
			%%Qualified: ArbWeb.Games.LearnMappings.WriteLearningsToFile
		----------------------------------------------------------------------------*/
		public void WriteLearningsToFile(string sSiteFileLTR, string sSiteFileRTL, string sTeamFileLTR, string sTeamFileRTL, string sNumberFileLTR, string sNumberFileRTL)
		{
			m_siteMapAndConfidences.WriteConfidencesToFile(sSiteFileLTR, sSiteFileRTL);
			m_teamMapAndConfidences.WriteConfidencesToFile(sTeamFileLTR, sTeamFileRTL);
			m_gameTagMapAndConfidences.WriteConfidencesToFile(sNumberFileLTR, sNumberFileRTL);
		}

		/*----------------------------------------------------------------------------
			%%Function: UpdateConfidencesForGames
			%%Qualified: ArbWeb.Games.LearnMappings.UpdateConfidencesForGames
		----------------------------------------------------------------------------*/
		public void UpdateConfidencesForGames(SimpleGame gameLeft, SimpleGame gameRight, int nConfidence)
		{
			m_siteMapAndConfidences.Update(gameLeft.Site, gameRight.Site, nConfidence);
			m_teamMapAndConfidences.Update(gameLeft.Home, gameRight.Home, nConfidence);
			m_teamMapAndConfidences.Update(gameLeft.Away, gameRight.Away, nConfidence);
			// only update this game mapping if we're almost certain, otherwise it will remain unmapped
			// (the game numbers match, or the day/date/time/slot fits)
			if (nConfidence >= 95)
				m_gameTagMapAndConfidences.Update(gameLeft.Number, gameRight.Number, nConfidence);
				
		}

		/*----------------------------------------------------------------------------
			%%Function: UpdatePossibleMaps
			%%Qualified: ArbWeb.Games.LearnMappings.UpdatePossibleMaps
		----------------------------------------------------------------------------*/
		public void UpdatePossibleMaps(SimpleGame game)
		{
			m_siteMapAndConfidences.AddPotentialMap(game.Site);
			m_teamMapAndConfidences.AddPotentialMap(game.Home);
			m_teamMapAndConfidences.AddPotentialMap(game.Away);
			m_gameTagMapAndConfidences.AddPotentialMap(game.Number);
		}
	}
}