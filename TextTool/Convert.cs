namespace Tools.TextTool
{
	public class Convert1
	{
		/// <summary>
		/// <, >, &를 &lt;, &gt;, &amp;로 변환
		/// </summary>
		public string RemoveHTMLtag(string str)
		{
			return str.Replace("<", "&lt;").Replace(">", "&gt;").Replace("&", "&amp;");
		}
	}
}
