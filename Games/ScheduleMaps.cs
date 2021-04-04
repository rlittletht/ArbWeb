using System.Collections.Generic;

namespace ArbWeb.Games
{
	// How to map team names and sites between two schedules
	public class ScheduleMaps
	{
		public Dictionary<string, string> TeamsMap { get; set; }
		public Dictionary<string, string> SitesMap { get; set; }

		public ScheduleMaps(LearnMappings.MapAndConfidences teamConfidences, LearnMappings.MapAndConfidences siteConfidences)
		{
			TeamsMap = teamConfidences.CreateMapFromLearning();
			SitesMap = siteConfidences.CreateMapFromLearning();
		}
	}
}