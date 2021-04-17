using System.Collections.Generic;

namespace ArbWeb.Games
{
	// this is a schedule, but with differences marked (inserted and deleted)
	public class SimpleDiffSchedule
	{
		// private SortedList<string, SimpleDiffGame> m_gamesByNumber = new SortedList<string, SimpleDiffGame>();
		private SortedList<string, SimpleDiffGame> m_games = new SortedList<string, SimpleDiffGame>();

		public IEnumerable<SimpleDiffGame> Games => m_games.Values;
		// public IEnumerable<SimpleDiffGame> GamesByNumber => m_gamesByNumber.Values;

		public void AddGame(SimpleDiffGame game)
		{
			// m_gamesByNumber.Add(game.Number, game);
			m_games.Add(game.SortKey, game);
		}
	}
}