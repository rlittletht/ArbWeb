using System;
using System.Collections.Generic;

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
				scheduleDiff.AddGame(new SimpleDiffGame(gameRightInTermsOfLeft, SimpleDiffGame.DiffOp.None));
			}
			else
			{
				scheduleDiff.AddGame(new SimpleDiffGame(gameLeft, SimpleDiffGame.DiffOp.Insert));
				scheduleDiff.AddGame(new SimpleDiffGame(gameRightInTermsOfLeft, SimpleDiffGame.DiffOp.Delete));
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
					? right.LookupGameNumber(maps.GameNumberMap[gameLeft.Number])
					: null;

				if (gameRight == null)
				{
					scheduleDiff.AddGame(new SimpleDiffGame(gameLeft, SimpleDiffGame.DiffOp.Delete));
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

			foreach (SimpleGame gameRight in right.Games)
			{
				for (int i = gamesMissingFromRight.Count - 1; i >= 0; i--)
				{
					SimpleGame gameMissing = gamesMissingFromRight[i];

					// if over 90% certain, we match!
					if (FuzzyMatcher.IsGameFuzzyMatch(gameMissing, gameRight) > 90)
					{
						if (hashGamesUsedFromRight.Contains(gameRight))
							throw new Exception("game we thought was missing was already matched in the diff");

						hashGamesUsedFromRight.Add(gameRight);
						maps.AddGameNumberMap(gameMissing.Number, gameRight.Number);
						RecordGameDiff(scheduleDiff, maps, null, gameRight);

						gamesMissingFromRight.RemoveAt(i);
					}
				}
			}

			// ok, what we have left are games that we couldn't match by number...
			if (gamesMissingFromRight.Count > 0)
				throw new Exception("Fuzzy non game number match NYI");
			
			foreach (SimpleGame gameRight in right.Games)
			{

			}

			return scheduleDiff;
		}
	}
}