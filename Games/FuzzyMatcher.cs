using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;

namespace ArbWeb.Games
{
	public class FuzzyMatcher
	{
		/*----------------------------------------------------------------------------
			%%Function: IsGameFuzzyMatch
			%%Qualified: ArbWeb.Games.LearnMappings.IsGameFuzzyMatch
		----------------------------------------------------------------------------*/
		public static int IsGameFuzzyMatch(SimpleGame gameLeft, SimpleGame gameRight)
		{
			int nConfidenceFuzzySiteMatch = FuzzyMatcher.IsGameFuzzySiteMatch(gameLeft, gameRight);
			int nConfidenceFuzzyTeamMatch = IsGameFuzzyTeamsMatch(gameLeft, gameRight);
			int nConfidenceFuzzyLevelMatch = IsGameFuzzyLevelMatch(gameLeft, gameRight);

			// if the site confidence is 0 AND the teams fuzzy match is 0, then even if the
			// game numbers "match", its a fail.
			if (nConfidenceFuzzySiteMatch == 0 && nConfidenceFuzzyTeamMatch == 0)
			{
				// we have no reason to believe this is the same game
				return 0;
			}

			// if the levels don't match, its not the same game
			if (nConfidenceFuzzyLevelMatch == 0)
				return 0;
			
			int numberLeft = Int32.Parse(gameLeft.Number);
			int numberRight = Int32.Parse(gameRight.Number);

			// if the numbers are an exact match, we only have to have low confidence
			// on the site matching to believe this is a certain match. otherwise, its just a 60%
			// match (all arbitrary numbers)
			if (numberLeft == numberRight)
				return (nConfidenceFuzzySiteMatch * nConfidenceFuzzyTeamMatch) / 10000 >= 30 ? 100 : Math.Max(95, nConfidenceFuzzySiteMatch);

			int dNum = Math.Abs(numberLeft - numberRight);

			if (dNum == (dNum / 1000) * 1000)
				return (nConfidenceFuzzySiteMatch * nConfidenceFuzzyTeamMatch) / 10000 >= 60 ? 100 : Math.Max(95, nConfidenceFuzzySiteMatch);

			// now a bunch of heuristics

			int nOverallConfidence = 0;

			nOverallConfidence = Math.Max(nOverallConfidence, (70 * nConfidenceFuzzyTeamMatch) / 100);
			nOverallConfidence = Math.Max(nOverallConfidence, (95 * nConfidenceFuzzySiteMatch) / 100);

			return nOverallConfidence;
		}

		enum FieldNumberMatchState
		{
			None,
			Mismatch,
			Match
		}

		/*----------------------------------------------------------------------------
			%%Function: IsStringFuzzySubstringMatch
			%%Qualified: ArbWeb.Games.FuzzyMatcher.IsStringFuzzySubstringMatch

			Return a 0-100 confidence value if these strings fuzzy match using parts
			of the string.
		----------------------------------------------------------------------------*/
		static int IsStringFuzzySubstringMatch(string sLeft, string sRight)
		{
			// easiest, exact match?
			if (string.Compare(sLeft, sRight) == 0)
				return 100;

			string sShorter = sLeft.Length < sRight.Length ? sLeft : sRight;
			string sLonger = sLeft.Length >= sRight.Length ? sLeft : sRight;

			// fully contained or substring?
			if (sLonger.Contains(sShorter))
				return 80;

			// every character in shorter is in longer, and in order?
			int iLast = 0;

			foreach (char ch in sShorter)
			{
				iLast = sLonger.IndexOf(ch, iLast);
				if (iLast == -1)
					break;
				iLast++; // don't let it keep rechecking the same character!
			}

			int nConfidenceLast = 0;

			if (iLast != -1)
				nConfidenceLast = (sShorter.Length > 5) ? 80 : 50;

			// ok, now we're going to try to do substring matches. numbers and directionality
			// cannot differ

			string[] rgsSubstrings = sShorter.Split(' ');
			int matchedLen = 0;
			FieldNumberMatchState fieldNumMatched = FieldNumberMatchState.None;

			foreach (string substring in rgsSubstrings)
			{
				if (substring.Length == 2 && substring[0] == '#')
				{
					// matching a field number
					if (sLonger.Contains(substring))
						fieldNumMatched = FieldNumberMatchState.Match;
					else
						fieldNumMatched = FieldNumberMatchState.Mismatch;

					// don't increase matched length for a matching field #
					continue;
				}
				
				if (sLonger.Contains(substring))
				{
					matchedLen += substring.Length;
					continue;
				}

				// no match. is this a disqualifying substring?
				if (string.Compare(substring, "NORTH") == 0
					|| string.Compare(substring, "SOUTH") == 0)
				{
					matchedLen = 0;
					break;
				}

				if (substring.Length == 1 && char.IsNumber(substring[0]))
				{
					matchedLen = 0;
					break;
				}

				if (substring.Length == 2
					&& substring[0] == '#')
				{
					if (!sLonger.Contains(substring[1]))
					{
						fieldNumMatched = FieldNumberMatchState.Mismatch;
						break;
					}
					else
					{
						fieldNumMatched = FieldNumberMatchState.Match;
						continue;
					}
				}
			}

			// if we had a field number and it mismatched, then it can't be
			// the same field
			if (fieldNumMatched == FieldNumberMatchState.Mismatch)
				return 0;
			
			if (matchedLen > 0)
			{
				// our confidence is based on how much of the substring was matched. 
				int nPctMatched = (matchedLen * 100) / sShorter.Length;

				if (nPctMatched > 90)
					nConfidenceLast = Math.Max(nConfidenceLast, 80);
				else if (nPctMatched > 50)
					nConfidenceLast = Math.Max(nConfidenceLast, 50);
				else if (nPctMatched > 15 && matchedLen > 1)
					nConfidenceLast = Math.Max(nConfidenceLast, 30);
			}

			return nConfidenceLast;
		}

		/*----------------------------------------------------------------------------
			%%Function: IsStringFuzzyMatch
			%%Qualified: ArbWeb.Games.LearnMappings.IsStringFuzzyMatch
		
			return confidence that these two strings fuzzy match (with special
			logic for aliasing different names for facilities like "park" or "field")
		----------------------------------------------------------------------------*/
		static int IsStringFuzzyMatchForSite(string sLeft, string sRight)
		{
			int nConfidence = 0;
			sLeft = sLeft.ToUpper();
			sRight = sRight.ToUpper();

			List<string> leftsToTry = new List<string>();

			leftsToTry.Add(sLeft);
			leftsToTry.Add(sLeft.Replace("PARK", "FIELD"));
			leftsToTry.Add(sLeft.Replace("PARK", "BALLFIELD"));
			leftsToTry.Add(sLeft.Replace("FIELD", "PARK"));
			leftsToTry.Add(sLeft.Replace("BALLFIELD", "PARK"));

			// first, try all the high confidence variations before we try lower confidence
			// fuzzy matches. this lets a variation (park->field) match on high confidence, even
			// if there was a lower confidence match without the variation
			foreach (string leftTry in leftsToTry)
			{
				nConfidence = Math.Max(nConfidence, IsStringFuzzySubstringMatch(leftTry, sRight));
			}

			return nConfidence;
		}

		/*----------------------------------------------------------------------------
			%%Function: IsGameFuzzySiteMatch
			%%Qualified: ArbWeb.Games.FuzzyMatcher.IsGameFuzzySiteMatch
		----------------------------------------------------------------------------*/
		public static int IsGameFuzzySiteMatch(SimpleGame gameLeft, SimpleGame gameRight)
		{
			return IsStringFuzzyMatchForSite(gameLeft.Site, gameRight.Site);
		}

		public static int IsGameFuzzyTeamsMatch(SimpleGame gameLeft, SimpleGame gameRight)
		{
			// first see if there's a match with Home/Home and Away/Away

			int nConfidenceBest = 0;

			int nConfidenceHome = IsStringFuzzySubstringMatch(gameLeft.Home, gameRight.Home);
			int nConfidenceAway = IsStringFuzzySubstringMatch(gameLeft.Away, gameRight.Away);

			nConfidenceBest = (nConfidenceHome * nConfidenceAway) / 100;

			nConfidenceHome = IsStringFuzzySubstringMatch(gameLeft.Home, gameRight.Away);
			nConfidenceAway = IsStringFuzzySubstringMatch(gameLeft.Away, gameRight.Home);

			nConfidenceBest = Math.Max(nConfidenceBest, (nConfidenceHome * nConfidenceAway) / 100);

			return nConfidenceBest;
		}

		enum Sport
		{
			Baseball,
			Softball,
			Unknown
		}

		/*----------------------------------------------------------------------------
			%%Function: SportFromSportString
			%%Qualified: ArbWeb.Games.FuzzyMatcher.SportFromSportString
		----------------------------------------------------------------------------*/
		static Sport SportFromSportString(string sSport)
		{
			if (string.IsNullOrEmpty(sSport))
				return Sport.Unknown;
			
			sSport = sSport.ToUpper();

			if (sSport.Contains("BB"))
				return Sport.Baseball;
			if (sSport.Contains("BASEBALL"))
				return Sport.Baseball;
			if (sSport.Contains("SB"))
				return Sport.Softball;
			if (sSport.Contains("SOFTBALL"))
				return Sport.Softball;
			return Sport.Unknown;
		}

		public static int IsGameFuzzySportMatch(SimpleGame gameLeft, SimpleGame gameRight)
		{
			Sport sportLeft = SportFromSportString(gameLeft.Sport);
			Sport sportRight = SportFromSportString(gameRight.Sport);

			if (sportLeft == Sport.Unknown || sportRight == Sport.Unknown)
				return 50;
			else if (sportLeft != sportRight)
				return 0;

			return 100;
		}
		public static int IsGameFuzzyLevelMatch(SimpleGame gameLeft, SimpleGame gameRight)
		{
			int adjust = IsGameFuzzySportMatch(gameLeft, gameRight);
			
			if (adjust == 0)
				return 0;
			
			return (IsStringFuzzySubstringMatch(gameLeft.Level.ToUpper(), gameRight.Level.ToUpper()) * adjust) / 100;
		}

		public enum Confidence
		{
			Certain,
			High,
			Medium,
			Low,
			None
		}

		[TestCase("foo", "foo", Confidence.Certain)]
		[TestCase("foo", "foobar", Confidence.High)]
		[TestCase("foo", "bar foo boo", Confidence.High)]
		[TestCase("BFH1", "big finn hill #1", Confidence.Medium)]
		[TestCase("H5", "hartman park #5", Confidence.Medium)]
		[TestCase("rrjr", "redmond ridge park big", Confidence.None)]
		[TestCase("Everest Field #1", "Everest Park #1", Confidence.Certain)]
		[TestCase("hidden valley #3", "hidden valley sports complex field #3", Confidence.High)]
		[TestCase("Big Rock 2", "Hartman Park #2", Confidence.None)]
		[TestCase("BFH #2", "East Sammamish Park #2", Confidence.None)]
		[TestCase("AAA", "Coast", Confidence.None)]
		[Test]
		public static void TestStringFuzzyMatch(string sLeft, string sRight, Confidence confidenceExpect)
		{
			int nConfidence = IsStringFuzzyMatchForSite(sLeft, sRight);
			Confidence confidenceActual;

			if (nConfidence == 100)
				confidenceActual = Confidence.Certain;
			else if (nConfidence >= 80)
				confidenceActual = Confidence.High;
			else if (nConfidence >= 50)
				confidenceActual = Confidence.Medium;
			else if (nConfidence > 0)
				confidenceActual = Confidence.Low;
			else
				confidenceActual = Confidence.None;

			Assert.AreEqual(confidenceExpect, confidenceActual);
		}
	}
}