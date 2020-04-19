using System;
using System.IO;
using System.Net;
using System.Text;
using Newtonsoft.Json;


namespace Tools.NetTool
{
	public class WebUtil
	{
		/// <summary>
		/// URL전송용 텍스트 인코딩
		/// </summary>
		/// <param name="EncType">EUC-KR, UTF-8</param>
		public static string encURL(string text, string EncType)
		{
			System.Text.Encoding enc = System.Text.Encoding.GetEncoding(EncType.ToUpper());
			byte[] bytesData = enc.GetBytes(text);
			return System.Web.HttpUtility.UrlEncode(bytesData, 0, bytesData.Length);
		}

		/// <summary>
		/// 특수기호를 변환(&lt;, &gt;, &quot;)
		/// </summary>
		public static string toHTML(string text)
		{
			return text.Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
		}

		/// <summary>
		/// 웹페이지 일부 화면 캡쳐, 저장된 파일 경로 반환
		/// </summary>
		/// <param name="URL"></param>
		/// <param name="filename">경로를 포함한 파일 이름</param>
		/// <param name="x">좌상단 X좌표</param>
		/// <param name="y">좌상단 Y좌표</param>
		public static string WebpageCapture(string URL, string filename, int x, int y, int width, int height, System.Drawing.Imaging.ImageFormat objImgFormat)
		{
			WebpageCapture objCapture = new WebpageCapture(URL, filename, x, y, width, height, objImgFormat);
			return objCapture.RunCapture();
		}

		/// <summary>
		/// 웹페이지 전체화면 캡쳐, 저장된 파일 경로 반환
		/// </summary>
		/// <param name="filename">경로를 포함한 파일 이름</param>
		public static string WebpageCapture(string URL, string filename, System.Drawing.Imaging.ImageFormat objImgFormat)
		{
			WebpageCapture objCapture = new WebpageCapture(URL, filename, objImgFormat);
			return objCapture.RunCapture();
		}

		/// <summary>
		/// 웹페이지 일부 화면 캡쳐
		/// </summary>
		/// <param name="filename">경로를 포함한 파일 이름</param>
		public static void WebpageCapture_noReturn(string URL, string filename, int x, int y, int width, int height, System.Drawing.Imaging.ImageFormat objImgFormat)
		{
			WebpageCapture objCapture = new WebpageCapture(URL, filename, x, y, width, height, objImgFormat);
			objCapture.WebpageCapture_noReturn();
		}

		/// <summary>
		/// 웹페이지 전체화면 캡쳐
		/// </summary>
		/// <param name="filename">경로를 포함한 파일 이름</param>
		public static void WebpageCapture_noReturn(string URL, string filename, System.Drawing.Imaging.ImageFormat objImgFormat)
		{
			WebpageCapture objCapture = new WebpageCapture(URL, filename, objImgFormat);
			objCapture.WebpageCapture_noReturn();
		}

		/// <summary>
		/// Google URL 줄임 서비스, 
		/// </summary>
		public static string URLshorten(string longurl)
		{
			WebRequest webreq;

			string jsonData;
			string rtnJson;
			byte[] bytebuffer;
			

			jsonData = JsonConvert.SerializeObject(new { longurl });
			bytebuffer = Encoding.UTF8.GetBytes(jsonData);//string json객체를 바이트코드로 변환

			webreq = WebRequest.Create("https://www.googleapis.com/urlshortener/v1/url");//google URL shorten API
			webreq.Method = WebRequestMethods.Http.Post;
			webreq.ContentType = "application/json";
			webreq.ContentLength = bytebuffer.Length;

			//리퀘스트
			using (Stream RequestStream = webreq.GetRequestStream())
			{
				RequestStream.Write(bytebuffer, 0, bytebuffer.Length);
				RequestStream.Close();
			}

			//리스폰스
			using (HttpWebResponse webresp = (HttpWebResponse)webreq.GetResponse())
			{
				using (Stream ResponseStream = webresp.GetResponseStream())
				{
					using (StreamReader sr = new StreamReader(ResponseStream))
					{
						rtnJson = sr.ReadToEnd();
					}
				}
			}

			dynamic objRcvJson = JsonConvert.DeserializeObject(rtnJson);

			return objRcvJson.id.ToString();
		}

		/// <summary>
		/// POST방식으로 데이터 전송
		/// </summary>
		/// <param name="timeOut">밀리세컨드</param>
		/// <param name="EncodingType">ex) EUC-KR, UTF-8</param>
		/// <returns>URL의 응답</returns>
		public static string SendPostData(string SendData, string URL, string EncodingType, int timeOut)
		{
			System.Diagnostics.Stopwatch objStopWatch = new System.Diagnostics.Stopwatch();

			HttpWebRequest httpWebRequest;
			HttpWebResponse httpWebResponse;
			Stream requestStream;
			StreamReader streamReader;
			byte[] Data;
			string Result;
			string encType = EncodingType.ToLower();

			if (!encType.Equals("euc-kr"))
				encType = "utf-8";//euc-kr이 아닌 기타 입력은 utf-8로 처리

			Data = System.Text.Encoding.GetEncoding(encType).GetBytes(SendData);

			try
			{
				objStopWatch.Start();
				httpWebRequest = (HttpWebRequest)WebRequest.Create(URL);
				httpWebRequest.ContentType = "application/x-www-form-urlencoded; charset=" + EncodingType;
				httpWebRequest.Method = "POST";
				httpWebRequest.ContentLength = Data.Length;

				//타임아웃 설정
				httpWebRequest.Timeout = timeOut;
				httpWebRequest.ReadWriteTimeout = timeOut;

				requestStream = httpWebRequest.GetRequestStream();
				requestStream.Write(Data, 0, Data.Length);
				requestStream.Close();

				httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
				streamReader = new StreamReader(httpWebResponse.GetResponseStream());

				Result = streamReader.ReadToEnd();
				streamReader.Close();
				httpWebResponse.Close();
			}
			catch (Exception ex)
			{
				objStopWatch.Stop();
				new TextTool.Util().writeLog($"FAIL SEND POST TRANSFER : {ex}\r\nELAPSED TIME:{objStopWatch.Elapsed}");
				Result = "FAIL";
			}

			if (objStopWatch.IsRunning)
				objStopWatch.Stop();

			return Result;
		}
		/// <summary>
		/// GET방식으로 데이터 전송
		/// </summary>
		/// <param name="timeOut">밀리세컨드</param>
		/// <returns>URL의 응답</returns>
		public static string SendQueryString(string SendData, string URL, int timeOut)
		{
			System.Diagnostics.Stopwatch objStopWatch = new System.Diagnostics.Stopwatch();
			MyWebClient WC = new MyWebClient(timeOut);

			string Result;

			try
			{
				objStopWatch.Start();
				Result = WC.DownloadString(URL + "?" + SendData);
			}
			catch (Exception ex)
			{
				objStopWatch.Stop();
				new TextTool.Util().writeLog($"FAIL SEND GET TRANSFER : {ex}\r\nELAPSED TIME:{objStopWatch.Elapsed}");
				Result = "FAIL";
			}

			if (objStopWatch.IsRunning)
				objStopWatch.Stop();

			return Result;
		}
		class MyWebClient : WebClient
		{
			protected int timeOut;
			public MyWebClient(int TimeOut)
			{
				this.timeOut = TimeOut;
			}
			protected override WebRequest GetWebRequest(Uri address)
			{
				WebRequest request = base.GetWebRequest(address);
				request.Timeout = this.timeOut;
				return request;
			}
		}
		/// <summary>
		/// iphone, android는 true
		/// </summary>
		public static bool isMobileDevice(System.Web.HttpRequest objRequest)
		{
			string userAgent = objRequest.UserAgent.ToLower();

			if (userAgent.IndexOf("iphone") > -1 || userAgent.IndexOf("android") > -1)
				return true;
			else
				return false;
		}
	}
}
