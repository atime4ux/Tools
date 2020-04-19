using System;
using System.Collections.Specialized;
using System.Data;
using System.Text;
using System.Web;
using System.Web.SessionState;

namespace Tools.NetTool
{
	public class WebFunc
	{
		public WebFunc()
		{
		}

		public string ChkBoxBuilder(DataTable DT, string conName, int HCount, string Attr)
		{
			string str;
			StringBuilder stringBuilder = new StringBuilder();
			if (DT.Columns.Count >= 2)
			{
				for (int i = 0; i < DT.Rows.Count; i++)
				{
					DataRow item = DT.Rows[i];
					if ((i + 1) % HCount == 0)
					{
						stringBuilder.Append("<br/>");
					}
					stringBuilder.Append("<input type=\"checkbox\" ");
					stringBuilder.Append(string.Concat(" name=\"", conName, "\" "));
					stringBuilder.Append(string.Concat(" id=\"", conName, "\" "));
					stringBuilder.Append(string.Concat(" value=\"", item[1].ToString(), "\" "));
					stringBuilder.Append(Attr);
					stringBuilder.Append(" >&nbsp;");
					stringBuilder.Append(item[0].ToString());
					stringBuilder.Append("&nbsp;\n");
				}
				str = stringBuilder.ToString();
			}
			else
			{
				str = "";
			}
			return str;
		}

		public string Form(HttpRequest RequestState, string FormName)
		{
			string str;
			int num = 0;
			while (true)
			{
				if (num >= RequestState.Form.Count)
				{
					str = "";
					break;
				}
				else if (!(RequestState.Form.GetKey(num) == FormName))
				{
					num++;
				}
				else
				{
					str = RequestState.Form.Get(num).ToString();
					break;
				}
			}
			return str;
		}

		public string GetCookie(HttpRequest RequestState, string CookieName)
		{
			string str;
			str = (RequestState.Cookies[CookieName] != null ? HttpUtility.HtmlEncode(RequestState.Cookies[CookieName].Value).ToString() : "");
			return str;
		}

		public string getDbString(string IN)
		{
			IN.Replace("'", "''");
			return IN;
		}

		public string getRequest(HttpRequest RequestState, string strName)
		{
			string str;
			string str1 = this.Form(RequestState, strName);
			str = (str1.Length <= 0 ? this.QueryString(RequestState, strName) : str1);
			return str;
		}

		public string getSession(HttpSessionState SessionState, string SessionName)
		{
			string str;
			int num = 0;
			while (true)
			{
				if (num >= SessionState.Keys.Count)
				{
					str = "";
					break;
				}
				else if (!(SessionState.Keys.Get(num).ToString() == SessionName))
				{
					num++;
				}
				else if (SessionState[num] != null)
				{
					str = SessionState[num].ToString();
					break;
				}
				else
				{
					str = "";
					break;
				}
			}
			return str;
		}

		public string getSession(HttpSessionState SessionState, int SessionIdx)
		{
			string str;
			str = (SessionState.Count - 1 >= SessionIdx ? SessionState[SessionIdx].ToString() : "");
			return str;
		}

		public string getUrlEncodeEuc(string val)
		{
			return HttpUtility.UrlEncode(val, Encoding.GetEncoding(949));
		}

		public string OptionBuilder(DataTable DT, string DefValue)
		{
			string str;
			StringBuilder stringBuilder = new StringBuilder();
			if (DT.Columns.Count >= 2)
			{
				for (int i = 0; i < DT.Rows.Count; i++)
				{
					DataRow item = DT.Rows[i];
					stringBuilder.Append(string.Concat("<option value = \"", item[1].ToString(), "\" "));
					if (item[1].ToString() == DefValue)
					{
						stringBuilder.Append(" Selected ");
					}
					stringBuilder.Append(" >");
					stringBuilder.Append(item[0].ToString());
					stringBuilder.Append("</option>\n");
				}
				str = stringBuilder.ToString();
			}
			else
			{
				str = "";
			}
			return str;
		}

		public string PrintState(HttpSessionState NowSession)
		{
			StringBuilder stringBuilder = new StringBuilder();
			for (int i = 0; i < NowSession.Keys.Count; i++)
			{
				stringBuilder.AppendLine(string.Concat(NowSession.Keys[i].ToString(), ":", NowSession[i].ToString()));
			}
			return stringBuilder.ToString();
		}

		public string QueryString(HttpRequest RequestState, string QueryName)
		{
			string str;
			int num = 0;
			while (true)
			{
				if (num >= RequestState.QueryString.Count)
				{
					str = "";
					break;
				}
				else if (!(RequestState.QueryString.GetKey(num) == QueryName))
				{
					num++;
				}
				else
				{
					str = RequestState.QueryString.Get(num).ToString();
					break;
				}
			}
			return str;
		}

		public void SetCookie(HttpResponse ResponseState, string CookieName, string CookieValue, int ExpDays)
		{
			ResponseState.Cookies[CookieName].Value = CookieValue;
			if (ExpDays > 0)
			{
				ResponseState.Cookies[CookieName].Expires = DateTime.Now.AddDays((double)ExpDays);
			}
		}

		public void setSession(HttpSessionState SessionState, string SessionName, string SessionValue)
		{
			int num = 0;
			while (true)
			{
				if (num >= SessionState.Keys.Count)
				{
					SessionState.Add(SessionName, SessionValue);
					break;
				}
				else if (!(SessionState.Keys.Get(num).ToString() == SessionName))
				{
					num++;
				}
				else
				{
					SessionState.RemoveAt(num);
					SessionState.Add(SessionName, SessionValue);
					break;
				}
			}
		}
	}
}