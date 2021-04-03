using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ArbWeb.Games
{
	// a schedule of simple games
	public class SimpleSchedule
	{
		private SortedList<string, SimpleGame> m_gamesByNumber = new SortedList<string, SimpleGame>();
		private SortedList<string, SimpleGame> m_games = new SortedList<string, SimpleGame>();

		public IEnumerable<SimpleGame> Games => m_games.Values;
		public IEnumerable<SimpleGame> GamesByNumber => m_gamesByNumber.Values;

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
			m_gamesByNumber.Add(game.Number, game);
			m_games.Add(game.SortKey, game);
		}
		
		
	}
}