using System.Collections.Generic;

namespace ArbWeb.Games
{
	// How to map team names and sites between two schedules
	public class ScheduleMaps
	{
		public Dictionary<string, string> TeamsMap { get; set; }
		public Dictionary<string, string> SitesMap { get; set; }
		public Dictionary<string, string> GameNumberMap { get; set; }
		public Dictionary<string, string> TeamsMapReverse { get; set; }
		public Dictionary<string, string> SitesMapReverse { get; set; }
		public Dictionary<string, string> GameNumberMapReverse { get; set; }
		// Also add a reverse game map so we can go right to left when diffing!

		public ScheduleMaps(
			LearnMappings.MapAndConfidences teamConfidences, 
			LearnMappings.MapAndConfidences siteConfidences, 
			LearnMappings.MapAndConfidences gameTagConfidences)
		{
			TeamsMap = teamConfidences.CreateMapFromLearning();
			SitesMap = siteConfidences.CreateMapFromLearning();
			GameNumberMap = gameTagConfidences.CreateMapFromLearning();

			TeamsMapReverse = teamConfidences.CreateReverseMapFromLearning();
			SitesMapReverse = siteConfidences.CreateReverseMapFromLearning();
			GameNumberMapReverse = gameTagConfidences.CreateReverseMapFromLearning();
		}

		/*----------------------------------------------------------------------------
			%%Function: CreateGameLeftFromRight
			%%Qualified: ArbWeb.Games.ScheduleMaps.CreateGameLeftFromRight
		----------------------------------------------------------------------------*/
		public SimpleGame CreateGameLeftFromRight(SimpleGame gameRight)
		{
			string siteLeft = SitesMapReverse.ContainsKey(gameRight.Site.ToUpper())
				? SitesMapReverse[gameRight.Site.ToUpper()]
				: $"##{gameRight.Site}";

			string homeLeft = TeamsMapReverse.ContainsKey(gameRight.Home.ToUpper())
				? TeamsMapReverse[gameRight.Home.ToUpper()]
				: $"##{gameRight.Home}";

			string awayLeft = TeamsMapReverse.ContainsKey(gameRight.Away.ToUpper())
				? TeamsMapReverse[gameRight.Away.ToUpper()]
				: $"##{gameRight.Away}";

			string numberLeft = GameNumberMapReverse.ContainsKey(gameRight.Number.ToUpper())
				? GameNumberMapReverse[gameRight.Number.ToUpper()]
				: $"##{gameRight.Number}";

			return new SimpleGame(gameRight.StartDateTime, siteLeft, gameRight.Level, homeLeft, awayLeft, numberLeft, gameRight.Status, gameRight.Sport);
		}

		/*----------------------------------------------------------------------------
			%%Function: AddGameNumberMap
			%%Qualified: ArbWeb.Games.ScheduleMaps.AddGameNumberMap
		----------------------------------------------------------------------------*/
		public void AddGameNumberMap(string numberLeft, string numberRight)
		{
			if (!GameNumberMap.ContainsKey(numberLeft))
				GameNumberMap.Add(numberLeft, numberRight);
			
			if (!GameNumberMapReverse.ContainsKey(numberRight))
				GameNumberMapReverse.Add(numberRight, numberLeft);
		}
	}
}