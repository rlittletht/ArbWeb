using System.Collections.Generic;

namespace ArbWeb.Games
{
	// ================================================================================
	//  G A M E
	//
	// A game is just a collection of GameSlot. Each GameSlot knows everything about
	// the game (yes, its duplicated on each slot)
	// ================================================================================
	public class Game
	{
		private int m_cSlots;
		private int m_cOpen;

		private List<GameSlot> gameSlots;

		public Game()
		{
			gameSlots = new List<GameSlot>();
		}

		public void AddGameSlot(GameSlot gms)
		{
			gameSlots.Add(gms);
			m_cSlots++;
			if (gms.Open)
				m_cOpen++;
		}

		public int OpenSlots
		{
			get { return m_cOpen; }
		}

		public int TotalSlots
		{
			get { return m_cSlots; }
		}

		public List<GameSlot> Slots
		{
			get { return gameSlots; }
		}
	}
}