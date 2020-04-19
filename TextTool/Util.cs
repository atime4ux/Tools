using System;
using System.Collections;
using System.Collections.Specialized;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace Tools.TextTool
{
	public class Util
	{
		public string getAppCfg(string Key)
		{
			return ConfigurationManager.AppSettings.Get(Key);
		}

		public string getDBValue(DataRow DR, int idx)
		{
			string str;
			str = (!DR.IsNull(idx) ? DR[idx].ToString() : "");
			return str;
		}

		public string getDBValue(DataRow DR, string column)
		{
			string str;
			str = (!DR.IsNull(column) ? DR[column].ToString() : "");
			return str;
		}

		public bool isNumber(string strNum)
		{
			bool flag;
			ArrayList arrayLists = new ArrayList();
			string str = "0123456789";
			int num = 0;
			while (true)
			{
				if (num >= arrayLists.Count)
				{
					flag = true;
					break;
				}
				else if (str.IndexOf(Convert.ToChar(arrayLists[num])) >= 0)
				{
					num++;
				}
				else
				{
					flag = false;
					break;
				}
			}
			return flag;
		}

		public string Left(string OrgString, int Length)
		{
			string str;
			str = (OrgString.Length > Length ? OrgString.Substring(0, Length) : OrgString);
			return str;
		}

		public string LeftB(string str, int max)
		{
			string str1;
			string str2 = "";
			if (!(str.ToString() == ""))
			{
				string str3 = str.Trim().ToString();
				int num = 0;
				char[] charArray = str3.ToCharArray();
				if (str3.Length != 0)
				{
					int num1 = 0;
					while (num1 < (int)charArray.Length)
					{
						int num2 = Convert.ToInt32(charArray[num1]);
						if ((num2 < 0 ? false : num2 < 128))
						{
							num++;
						}
						else
						{
							num += 2;
						}
						if (num > max)
						{
							break;
						}
						else
						{
							str2 = string.Concat(str2, str3.Substring(num1, 1));
							num1++;
						}
					}
				}
				str1 = str2;
			}
			else
			{
				str1 = str2;
			}
			return str1;
		}

		public string LeftDot(string OrgString, int Length)
		{
			string str;
			str = (OrgString.Length > Length + 2 ? string.Concat(OrgString.Substring(0, Length), "..") : OrgString);
			return str;
		}

		public string RemoveTag(string str)
		{
			string str1 = str.Replace("<", "&lt;").Replace(">", "&gt;").Replace("&", "&amp;").Replace("'", "''");
			return str1;
		}

		public string Right(string OrgString, int Length)
		{
			string str;
			str = (OrgString.Length >= Length ? OrgString.Substring(OrgString.Length - Length, Length) : OrgString);
			return str;
		}

		public string[] Split(string OrgString, string Spliter)
		{
			return OrgString.Split(Spliter.ToCharArray());
		}

		public string ToCurrency(object obj)
		{
			string str = string.Format("{0:N0}", this.ToDecimal(obj.ToString()));
			return str;
		}

		public DateTime ToDate(string str)
		{
			DateTime dateTime;
			try
			{
				dateTime = Convert.ToDateTime(str);
			}
			catch (Exception exception1)
			{
				Exception exception = exception1;
				this.writeLog("Conv", string.Concat("DATE CONVERTING ERR : ", exception.ToString()));
				dateTime = Convert.ToDateTime("1945-08-15 00:00:00");
			}
			return dateTime;
		}

		public string ToDb(object txt)
		{
			string str;
			string str1;
			try
			{
				str = (txt.ToString().ToLower().IndexOf("information_schema.") != -1 ? "" : txt.ToString().Replace("'", "''").Trim());
				str1 = str;
			}
			catch (Exception exception)
			{
				this.writeLog(exception.ToString());
				str1 = "";
			}
			return str1;
		}

		public decimal ToDecimal(object obj)
		{
			decimal num;
			try
			{
				num = Convert.ToDecimal(obj);
			}
			catch (Exception exception1)
			{
				Exception exception = exception1;
				this.writeLog("Conv", string.Concat("DECIMAL CONVERTING ERR : ", exception.ToString()));
				this.writeLog("Conv", string.Concat("DECIMAL CONVERTING ERR OBJ: ", obj.ToString()));
				num = new decimal(0);
			}
			return num;
		}

		public double ToDouble(object obj)
		{
			double num;
			try
			{
				num = Convert.ToDouble(obj);
			}
			catch (Exception exception1)
			{
				Exception exception = exception1;
				this.writeLog("Conv", string.Concat("DOUBLE CONVERTING ERR : ", exception.ToString()));
				this.writeLog("Conv", string.Concat("DOUBLE CONVERTING ERR OBJ: ", obj.ToString()));
				num = 0;
			}
			return num;
		}

		public string ToFloatCurrency(object obj, int fsize)
		{
			string str = string.Format(string.Concat("{0:N", fsize.ToString(), "}"), this.ToDouble(obj.ToString()));
			return str;
		}

		public string ToFloatString(object obj, int fsize)
		{
			string str = string.Format(string.Concat("{0:F", fsize.ToString(), "}"), this.ToDouble(obj.ToString()));
			return str;
		}

		public string ToHtml(string org)
		{
			string str = org.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace(" ", "&nbsp;").Replace("\n", "<br/>\n");
			return str;
		}

		public int ToInt32(object obj)
		{
			int num;
			try
			{
				num = Convert.ToInt32(obj);
			}
			catch (Exception exception1)
			{
				Exception exception = exception1;
				this.writeLog("Conv", string.Concat("INT CONVERTING ERR : ", exception.ToString()));
				this.writeLog("Conv", string.Concat("INT CONVERTING ERR OBJ: ", obj.ToString()));
				num = 0;
			}
			return num;
		}

		public string ToJavascript(string str)
		{
			string str1 = str.Replace("\\", "\\\\").Replace("'", "\\'");
			return str1;
		}

		public long ToLong(object obj)
		{
			long num;
			try
			{
				num = Convert.ToInt64(obj);
			}
			catch (Exception exception1)
			{
				Exception exception = exception1;
				this.writeLog("Conv", string.Concat("LONG CONVERTING ERR : ", exception.ToString()));
				this.writeLog("Conv", string.Concat("LONG CONVERTING ERR : ", obj.ToString()));
				num = (long)0;
			}
			return num;
		}

		public void writeDebug(object txtLog)
		{
			try
			{
				this.writeLog("DEBUG", txtLog);
			}
			catch (Exception exception)
			{
				System.Diagnostics.Debug.WriteLine(exception.ToString());
			}
		}

		/// <summary>
		/// 비동기 처리할것
		/// </summary>
		/// <param name="txtLog"></param>
		public void writeLog(object txtLog)
		{
			try
			{
				DateTime now = DateTime.Now;
				string str = string.Concat("[", now.ToString("yyyy-MM-dd"), "]");
				string str1 = string.Concat(ConfigurationManager.AppSettings["logFile"], str, ".txt");
				FileStream fileStream = new FileStream(str1, FileMode.Append);
				StreamWriter streamWriter = new StreamWriter(fileStream, Encoding.UTF8);
				now = DateTime.Now;
				streamWriter.WriteLine(string.Concat("[", now.ToString("yyyy-MM-dd hh:mm:ss"), "]", txtLog.ToString()));
				streamWriter.WriteLine("");
				streamWriter.Flush();
				streamWriter.Close();
				fileStream.Close();
			}
			catch (Exception exception)
			{
				System.Diagnostics.Debug.WriteLine(exception.ToString());
			}
		}

		/// <summary>
		/// 비동기 처리할것
		/// </summary>
		/// <param name="Type"></param>
		/// <param name="txtLog"></param>
		public void writeLog(string Type, object txtLog)
		{
			try
			{
				DateTime now = DateTime.Now;
				string str = string.Concat("[", now.ToString("yyyy-MM-dd"), "]");
				string[] item = new string[] { ConfigurationManager.AppSettings["logFile"], str, "_", Type, ".txt" };
				FileStream fileStream = new FileStream(string.Concat(item), FileMode.Append);
				StreamWriter streamWriter = new StreamWriter(fileStream, Encoding.UTF8);
				now = DateTime.Now;
				streamWriter.WriteLine(string.Concat("[", now.ToString("yyyy-MM-dd hh:mm:ss"), "]", txtLog.ToString()));
				streamWriter.WriteLine("");
				streamWriter.Flush();
				streamWriter.Close();
				fileStream.Close();
			}
			catch (Exception exception)
			{
				System.Diagnostics.Debug.WriteLine(exception.ToString());
			}
		}
	}
}