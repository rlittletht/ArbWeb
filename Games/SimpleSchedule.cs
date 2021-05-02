using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ArbWeb.Games
{
	// a schedule of simple games
	public class SimpleSchedule
	{
		private SortedList<string, List<SimpleGame>> m_gamesByNumber = new SortedList<string, List<SimpleGame>>();
		private SortedList<string, SimpleGame> m_games = new SortedList<string, SimpleGame>();

		public IEnumerable<SimpleGame> Games => m_games.Values;
		public IEnumerable<IEnumerable<SimpleGame>> GamesByNumber => m_gamesByNumber.Values;

		/*----------------------------------------------------------------------------
			%%Function: LookupGameNumber
			%%Qualified: ArbWeb.Games.SimpleSchedule.LookupGameNumber
		----------------------------------------------------------------------------*/
		public SimpleGame FindGameByNumber(string gameNumber, SimpleGame gameMatch, HashSet<SimpleGame> gamesNotToConsider)
		{
			if (!m_gamesByNumber.ContainsKey(gameNumber))
				return null;

			SimpleGame gameBest = null;
			int nConfidenceBest = 0;
			
			foreach (SimpleGame gameCheck in m_gamesByNumber[gameNumber])
			{
				if (gamesNotToConsider != null && gamesNotToConsider.Contains(gameCheck))
					continue;
				
				int nConfidence = FuzzyMatcher.IsGameFuzzyMatch(gameMatch, gameCheck);

				if (nConfidence > nConfidenceBest)
				{
					gameBest = gameCheck;
					nConfidenceBest = nConfidence;
				}
			}

			return gameBest;
		}
		
		/*----------------------------------------------------------------------------
			%%Function:BuildFromScheduleGames
			%%Qualified:ArbWeb.Games.SimpleSchedule.BuildFromScheduleGames
		----------------------------------------------------------------------------*/
		public static SimpleSchedule BuildFromScheduleGames(ScheduleGames games)
		{
			SimpleSchedule schedule = new SimpleSchedule();

			HashSet<string> gamesSeen = new HashSet<string>();

			foreach (GameSlot gm in games.SortedGameSlots)
			{
				// only report each game once...
				if (gamesSeen.Contains(gm.GameNum))
					continue;

				gamesSeen.Add(gm.GameNum);
				schedule.AddSimpleGame(new SimpleGame(gm));
			}

			return schedule;
		}

		/*----------------------------------------------------------------------------
			%%Function:AddSimpleGame
			%%Qualified:ArbWeb.Games.SimpleSchedule.AddSimpleGame
		----------------------------------------------------------------------------*/
		public void AddSimpleGame(SimpleGame game)
		{
			if (!m_gamesByNumber.ContainsKey(game.Number))
				m_gamesByNumber.Add(game.Number, new List<SimpleGame>());
			
			m_gamesByNumber[game.Number].Add(game);
			m_games.Add(game.SortKey, game);
		}
	}
}