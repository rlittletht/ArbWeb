using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;

namespace ArbWeb.Games
{
	public class FuzzyMatcher
	{
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
			}

			int nConfidenceLast = 0;

			if (iLast != -1)
				nConfidenceLast = (sShorter.Length > 5) ? 80 : 50;

			// ok, now we're going to try to do substring matches. numbers and directionality
			// cannot differ

			string[] rgsSubstrings = sShorter.Split(' ');
			int matchedLen = 0;
			foreach (string substring in rgsSubstrings)
			{
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
						matchedLen = 0;
						break;
					}
					else
					{
						matchedLen++;
					}
				}
			}

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
		public static int IsStringFuzzyMatch(string sLeft, string sRight)
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
		[Test]
		public static void TestStringFuzzyMatch(string sLeft, string sRight, Confidence confidenceExpect)
		{
			int nConfidence = IsStringFuzzyMatch(sLeft, sRight);
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