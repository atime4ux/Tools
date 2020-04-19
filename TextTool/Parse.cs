using System;
using System.Collections.Generic;
using System.Text;

namespace Tools.TextTool
{
	public static class Parse
	{
		/// <summary>
		/// 성별 추출
		/// </summary>
		public static string parseGender(string Text)
		{
			string strRex;
			string Gender;

			strRex = "(남|여)(자|성)?";

			Gender = parseReg(Text, strRex);

			//성별 통일
			if (Gender.IndexOf("남") == 0)
			{
				Gender = "남자";
			}
			else if (Gender.IndexOf("여") == 0)
			{
				Gender = "여자";
			}

			return Gender;
		}

		/// <summary>
		/// 한글만 추출
		/// </summary>
		public static string parseKorean(string Text)
		{
			string strRex;

			strRex = "[가-힣]+";

			return parseReg(Text, strRex);
		}

		/// <summary>
		/// 숫자 또는 숫자와 문자 조합 추출
		/// </summary>
		/// <param name="min">숫자 최소 길이</param>
		/// <param name="max">숫자 최대 길이</param>
		/// <param name="suffix">숫자와 결합된 문자열들, '|'로 구분</param>
		public static string parseNum_KOR(string Text, int min, int max, string suffix)
		{
			string strRex;

			strRex = "[0-9]{" + min + "," + max + "}(" + suffix + ")?";

			return parseReg(Text, strRex);
		}

		/// <summary>
		/// 숫자와 문자의 조함 추출
		/// </summary>
		/// <param name="min">숫자 최소 길이</param>
		/// <param name="max">숫자 최대 길이</param>
		/// <param name="suffix">숫자와 결합된 문자열들, '|'로 구분</param>
		public static string parseNum_AND_KOR(string Text, int min, int max, string suffix)
		{
			string strRex;

			strRex = "[0-9]{" + min + "," + max + "}(" + suffix + ")";

			return parseReg(Text, strRex);
		}

		/// <summary>
		/// 숫자만 추출
		/// </summary>
		/// <param name="min">숫자 최소 길이</param>
		/// <param name="max">숫자 최대 길이</param>
		public static string parseNum(string Text, int min, int max)
		{
			string strRex;
			strRex = "[0-9]{" + min + "," + max + "}";
			return parseReg(Text, strRex);
		}

		/// <summary>
		/// 메시지에 정규식에 포함하는 문자열이 있으면 반환
		/// </summary>
		private static string parseReg(string Text, string Reg)
		{
			System.Text.RegularExpressions.Regex REx = new System.Text.RegularExpressions.Regex(Reg);

			return REx.Match(Text).Value;
		}
	}
}
