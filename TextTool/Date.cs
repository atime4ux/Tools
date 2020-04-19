using System;

namespace Tools.TextTool
{
	public class DateBetween
	{
		public DateTime From { get; set; }
		public DateTime To { get; set; }

		public DateBetween(DateTime from, DateTime to, bool ignoreTime = false)
		{
			if (ignoreTime == true)
			{
				Init(from.Year, from.Month, from.Day, to.Year, to.Month, to.Day);
			}
			else
			{
				From = from;
				To = to;
			}
		}
		public DateBetween(int yearFrom, int monthFrom, int yearTo, int monthTo)
		{
			Init(yearFrom, monthFrom, yearTo, monthTo);
		}
		public DateBetween(int yearFrom, int monthFrom, int dayFrom, int yearTo, int monthTo, int dayTo)
		{
			Init(yearFrom, monthFrom, dayFrom, yearTo, monthTo, dayTo);
		}
		
		private void Init(int yearFrom, int monthFrom, int yearTo, int monthTo)
		{
			Init(yearFrom, monthFrom, 1, yearTo, monthTo, DateTime.DaysInMonth(yearTo, monthTo));
		}

		private void Init(int yearFrom, int monthFrom, int dayFrom, int yearTo, int monthTo, int dayTo)
		{
			From = new DateTime(yearFrom, monthFrom, dayFrom);
			To = new DateTime(yearTo, monthTo, dayTo).AddDays(1).AddMilliseconds(-1);
		}

		public static DateBetween GetFullYear(int year)
		{
			return new DateBetween(year, 1, year, 12);
		}

		public static DateBetween GetFullMonth(int year, int month)
		{
			return new DateBetween(year, month, year, month);
		}

		public static DateBetween GetFullDay(int year, int month, int day)
		{
			return new DateBetween(year, month, day, year, month, day);
		}

		public static DateBetween GetYearToYear(int yearFrom, int yearTo)
		{
			return new DateBetween(yearFrom, 1, yearTo, 12);
		}

		public static DateBetween GetCurrentFullYear()
		{
			return GetFullYear(DateTime.Now.Year);
		}

		public static DateBetween GetCurrentFullMonth()
		{
			DateTime dateNow = DateTime.Now;
			return GetFullMonth(dateNow.Year, dateNow.Month);
		}

		public static DateBetween GetCurrentFullDay()
		{
			DateTime dateNow = DateTime.Now;
			return GetFullDay(dateNow.Year, dateNow.Month, dateNow.Day);
		}
	}
}
