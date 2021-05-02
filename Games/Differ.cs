using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace ArbWeb.Games
{
	// this will diff two simple schedules
	
	// The left schedule is assumed to be a subset of the right. That's easy if you can
	// trust the game number.
	
	// UNDONE: If a game is cancelled and rescheduled, and gets a new game number on the
	// right, we don't yet find this game. We could do this by knowing the team matchups
	// that are on the left but missing from the right (should be a small set), and then
	// find those same matchups on the right that have no match on the left. Any matches
	// are those missing games.
	
	// for cancelled or rained out games, we have a couple of complications:
	// 1) the TW schedule might have the same game number more than once (for rainout resched)
	//     (and the 2nd TW game might map to nothing in arbiter, because it wasn't rescheduled)
	//     (OR the 2nd TW game might map to a 2nd arbiter game, but with a DIFFERENT number)
	// 2) the TW schedule might have a game not cancelled, and arbiter marked cancelled
	//    (this needs to be treated as a game missing on the right)
	// 3) the TW schedule has a game marked cancelled, and arbiter has same game cancelled
	//    (so the numbers are matched, but this is still a game missing on the right
	//    (though it might not be there if it hasn't been rescheduled yet)
	
	public class Differ
	{
		/*----------------------------------------------------------------------------
			%%Function: RecordGameDiff
			%%Qualified: ArbWeb.Games.Differ.RecordGameDiff
		
			record a game in the diff schedule. if they are equal, then just add
			the game. otherwise, add as a remove/insert/both as appropriate.
		
			The maps allow us to compare games in two different "namespaces"
			(trainwreck or arbiter). Make sure we always add in the left namespace
			(so the diff uses a consistent namespace)
		----------------------------------------------------------------------------*/
		public static void RecordGameDiff(
			SimpleDiffSchedule scheduleDiff, 
			ScheduleMaps maps, 
			SimpleGame gameLeft, 
			SimpleGame gameRight)
		{
			SimpleGame gameRightInTermsOfLeft = maps.CreateGameLeftFromRight(gameRight);
			
			if (gameLeft == null)
			{
				scheduleDiff.AddGame(new SimpleDiffGame(gameRightInTermsOfLeft, SimpleDiffGame.DiffOp.Insert));
			}
			else if (gameLeft.IsEqual(gameRight, maps))
			{
				scheduleDiff.AddGame(new SimpleDiffGame(gameLeft, SimpleDiffGame.DiffOp.None));
			}
			else
			{
				scheduleDiff.AddGame(new SimpleDiffGame(gameLeft, SimpleDiffGame.DiffOp.Delete));
				scheduleDiff.AddGame(new SimpleDiffGame(gameRightInTermsOfLeft, SimpleDiffGame.DiffOp.Insert));
			}
		}
		
		/*----------------------------------------------------------------------------
			%%Function: BuildDiffFromSchedules
			%%Qualified: ArbWeb.Games.Differ.BuildDiffFromSchedules

			This will look at the left schedule and determine what is different
			in the right schedule. The scenario is a team schedule on the left, and the
			entire arbiter schedule on the right. There are necessarily A LOT of games
			on the right that aren't on the left.
		
			We are not concerned with things in the right schedule but missing from left.
			Left is considered the "truth".

			HOWEVER, we DO have to find games that moved. So when we think we have a
			game on the left that has been deleted, we will remember that and go
			looking for it on the right.
		----------------------------------------------------------------------------*/
		public static SimpleDiffSchedule BuildDiffFromSchedules(SimpleSchedule left, SimpleSchedule right)
		{
			// first, figure out how to map team and field names from left to right
			ScheduleMaps maps = LearnMappings.GenerateMapsFromSchedules(left, right);
			SimpleDiffSchedule scheduleDiff = new SimpleDiffSchedule();

			List<SimpleGame> gamesMissingFromRight = new List<SimpleGame>();
			HashSet<SimpleGame> hashGamesUsedFromRight = new HashSet<SimpleGame>();
			
			foreach (SimpleGame gameLeft in left.Games)
			{
				SimpleGame gameRight = maps.GameNumberMap.ContainsKey(gameLeft.Number)
					? right.FindGameByNumber(maps.GameNumberMap[gameLeft.Number], gameLeft, hashGamesUsedFromRight)
					: null;

				if (gameRight == null)
				{
					// remember this game for later, so we can try to find if it moved somewhere
					gamesMissingFromRight.Add(gameLeft);
				}
				else
				{
					hashGamesUsedFromRight.Add(gameRight);
					maps.AddGameNumberMap(gameLeft.Number, gameRight.Number);
					RecordGameDiff(scheduleDiff, maps, gameLeft, gameRight);
				}
			}

			// now we have a set of games that we couldn't match in the right. since we only
			// looked for games at the same date/time, this isn't surprising -- we miss all
			// moves. now, let's iterate the right schedule and look for games that are missing
			// from the left. there will be A LOT of them. Then let's try to fuzzy match to see
			// if its the same game (game numbers are your friend here!)

			// it would be nice to just go through the missing game list and match by 
			// number, but we don't have that reverse map built

			// sadly this will be n^2, but hopefully the missing game list is short!

			// if over 90% certain, we match!

			int nThreshold = 90;

			// first, find the matches we are almost certain about.
			// then do it again for lower thresholds
			while (nThreshold > 0 && gamesMissingFromRight.Count > 0)
			{
				foreach (SimpleGame gameRight in right.Games)
				{
					for (int i = gamesMissingFromRight.Count - 1; i >= 0; i--)
					{
						SimpleGame gameMissing = gamesMissingFromRight[i];

						if (FuzzyMatcher.IsGameFuzzyMatch(gameMissing, gameRight) > nThreshold)
						{
							if (hashGamesUsedFromRight.Contains(gameRight))
								continue; // we already matched with this one...gotta find another

							hashGamesUsedFromRight.Add(gameRight);
							maps.AddGameNumberMap(gameMissing.Number, gameRight.Number);
							RecordGameDiff(scheduleDiff, maps, gameMissing, gameRight);

							gamesMissingFromRight.RemoveAt(i);
						}
					}
				}

				nThreshold -= 20;
			}

			// ok, what we have left are games that we couldn't match by number...

			if (gamesMissingFromRight.Count > 0)
				MessageBox.Show($"Still have {gamesMissingFromRight.Count} missing from Arbiter schedule.");
			
			foreach (SimpleGame gameMissing in gamesMissingFromRight)
				scheduleDiff.AddGame(new SimpleDiffGame(gameMissing, SimpleDiffGame.DiffOp.Delete));

			return scheduleDiff;
		}
	}
}