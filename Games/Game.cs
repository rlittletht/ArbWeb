using System.Collections.Generic;

namespace ArbWeb.Games
{
	// ================================================================================
	//  G A M E  D A T A 
	// ================================================================================
	public class Game
	{
		private int m_cSlots;
		private int m_cOpen;

		private List<GameSlot> m_plgms;

		public Game()
		{
			m_plgms = new List<GameSlot>();
		}

		public void AddGameSlot(GameSlot gms)
		{
			m_plgms.Add(gms);
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
			get { return m_plgms; }
		}
	}
}