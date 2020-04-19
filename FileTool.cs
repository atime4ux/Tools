using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Drawing;

using System.IO;

namespace Tools
{
	public static class FileTool
	{
		/// <summary>
		/// jpg, jpeg, bmp, png 파일 확장자 판별
		/// </summary>
		/// <param name="fileName"></param>
		/// <returns></returns>
		public static bool isIMG_extension(string fileName)
		{
			string fileExtension;
			string[] IMG_extension = { ".jpg", ".jpeg", ".bmp", ".png", ".gif" };
			int i;

			try
			{
				fileExtension = System.IO.Path.GetExtension(fileName).ToLower();
			}
			catch (Exception ex)
			{
				new TextTool.Util().writeLog("ERR GETTING FILE EXE : " + fileName + "\r\n" + ex.ToString());
				fileExtension = "";
			}

			for (i = 0; i < IMG_extension.Length; i++)
			{
				if (fileExtension.Equals(IMG_extension[i]))
				{
					return true;
				}
			}

			return false;
		}

		/// <summary>
		/// swf, avi, wmv, k3g 파일 확장자 판별
		/// </summary>
		/// <param name="fileName"></param>
		/// <returns></returns>
		public static bool isMOV_extension(string fileName)
		{
			string fileExtension;
			string[] MOV_extension = { ".swf" };
			int i;

			try
			{
				fileExtension = System.IO.Path.GetExtension(fileName).ToLower();
			}
			catch (Exception ex)
			{
				new TextTool.Util().writeLog("ERR GETTING FILE EXE : " + fileName + "\r\n" + ex.ToString());
				fileExtension = "";
			}

			for (i = 0; i < MOV_extension.Length; i++)
			{
				if (fileExtension.Equals(MOV_extension[i]))
				{
					return true;
				}
			}

			return false;
		}


		/// <summary>
		/// jpg가 아닌 이미지 파일을 jpg로 저장후 jpg파일 경로 리턴
		/// </summary>
		public static string SaveJPG(string Path)
		{
			string fileName = "";
			string fileExtension;

			fileExtension = System.IO.Path.GetExtension(Path);

			if (!fileExtension.Equals(".jpg"))
			{
				try
				{
					Image orgImg = Image.FromFile(Path);

					fileName = Path.Substring(0, Path.LastIndexOf(".")) + ".jpg";

					orgImg.Save(fileName, System.Drawing.Imaging.ImageFormat.Jpeg);
				}
				catch (Exception e)
				{
					new TextTool.Util().writeLog("SAVE TO JPG FAIL : " + e.ToString());
				}
			}
			else
			{
				fileName = Path;
			}

			return fileName;
		}

		/// <summary>
		/// 썸네일 이미지 생성후 파일 경로 리턴
		/// </summary>
		/// <param name="imgWidth">썸네일 이미지 가로 픽셀</param>
		public static string MakeThumbnail(string Path, int imgWidth)
		{
			string fileName = "";

			int width = imgWidth;
			int height;

			try
			{
				Image orgImg = Image.FromFile(Path);
				fileName = Path.Substring(0, Path.LastIndexOf(".")) + "_tmb.jpg";

				height = orgImg.Height * width / orgImg.Width;

				if (orgImg.Width < orgImg.Height)
				{
					//세로 사진
					if (width == 0) width = 160;
					if (height == 0) height = 212;
				}
				else
				{
					//가로 사진
					if (width == 0) width = 160;
					if (height == 0) height = 120;
				}

				System.Drawing.Image thmbImg = orgImg.GetThumbnailImage(width, height, delegate { return false; }, IntPtr.Zero);

				thmbImg.Save(fileName, System.Drawing.Imaging.ImageFormat.Jpeg);
			}
			catch (Exception e)
			{
				new TextTool.Util().writeLog("THUMBNAIL FAIL : " + e.ToString());
			}

			return fileName;
		}

		/// <summary>
		/// url의 "이미지" 파일을 다운받아서 저장, 저장된 서버주소 리턴
		/// </summary>
		/// <param name="url">다운 받을 파일의 url</param>
		/// <param name="urlBase">외부에서 저장된 파일에 액세스할 주소</param>
		/// <param name="downPath">저장할 로컬 경로</param>
		public static string FileSave(string url, string urlBase, string downPath)
		{
			string retUrl = "";
			//저장될 폴더
			string folder;
			//저장될 파일이름
			string fileName;
			//파일이름을 포함한 저장경로
			string fullpath_source;
			//파일이름을 포함한 저장경로(jpg로 변환 후)
			string fullpath_jpg;
			//확장자
			string fileExtension;

			try
			{
				//디렉토리 이름
				folder = DateTime.Now.ToString("yyyyMMdd");
				//경로 잘라내고 파일이름만 추출
				fileName = GetFileNameFromUrl(url);
				//확장자 추출
				fileExtension = System.IO.Path.GetExtension(fileName);
				//jpeg확장자는 jpg로 저장한다
				if (fileExtension.Equals(".jpeg"))
				{
					fileName = fileName.Substring(0, fileName.LastIndexOf(".")) + ".jpg";
				}

				//파일경로와 이름을 포함한 전체 경로 생성
				//중복되는 파일 존재할때 카운터 증가로 재생성
				fullpath_source = GetUniqueFileName(downPath + folder + "\\", fileName, true);

				//디렉토리 생성
				MakeDir(downPath + folder + "\\");
				//url에서 경로로 파일 다운로드
				if (DownloadFromUrl(url, fullpath_source))
				{
					//해당 파일의 썸네일 생성
					MakeThumbnail(fullpath_source, 160);
					//jpg로 변환
					fullpath_jpg = SaveJPG(fullpath_source);
					//로컬 주소를 웹주소로 변환 후 리턴
					if (!urlBase.Substring(urlBase.Length - 1).Equals("/"))
					{
						urlBase = urlBase + "/";
					}
					retUrl = urlBase + folder + "/" + GetFileNameFromUrl(fullpath_jpg);
				}
			}
			catch (Exception ex)
			{
				new TextTool.Util().writeLog("FileSave ERR URL : " + url);
				new TextTool.Util().writeLog("FileSave ERR : " + ex.ToString());
			}

			return retUrl;
		}

		/// <summary>
		/// url의 파일을 다운받아서 저장, 저장된 서버주소 리턴
		/// </summary>
		/// <param name="url">다운 받을 파일의 url</param>
		/// <param name="urlBase">외부에서 저장된 파일에 액세스할 주소</param>
		/// <param name="downPath">저장할 로컬 경로</param>
		public static string FileSave_NotIMG(string url, string urlBase, string downPath)
		{
			string retUrl = "";
			//저장될 폴더
			string folder;
			//저장될 파일이름
			string fileName;
			//파일이름을 포함한 저장경로
			string fullpath_source;
			//확장자
			string fileExtension;

			try
			{
				//디렉토리 이름
				folder = DateTime.Now.ToString("yyyyMMdd");
				//경로 잘라내고 파일이름만 추출
				fileName = GetFileNameFromUrl(url);
				//확장자 추출
				fileExtension = System.IO.Path.GetExtension(fileName);

				//파일경로와 이름을 포함한 전체 경로 생성
				//중복되는 파일 존재할때 카운터 증가로 재생성
				fullpath_source = GetUniqueFileName(downPath + folder + "\\", fileName, true);

				//디렉토리 생성
				MakeDir(downPath + folder + "\\");
				//url에서 경로로 파일 다운로드
				if (DownloadFromUrl(url, fullpath_source))
				{
					//로컬 주소를 웹주소로 변환 후 리턴
					if (!urlBase.Substring(urlBase.Length - 1).Equals("/"))
					{
						urlBase = urlBase + "/";
					}
					retUrl = urlBase + folder + "/" + GetFileNameFromUrl(fullpath_source);
				}
			}
			catch (Exception ex)
			{
				new TextTool.Util().writeLog("MovFileSave ERR URL :" + url);
				new TextTool.Util().writeLog("MovFileSave ERR :" + ex.ToString());
			}

			return retUrl;
		}

		/// <summary>
		/// url에서 파일이름만 리턴
		/// </summary>
		public static string GetFileNameFromUrl(string url)
		{
			url = url.Replace(@"\\", "/");
			url = url.Replace(@"\", "/");

			string[] splited = url.Split('/');
			if (splited == null || splited.Length == 0)
				return "";
			else
				return splited[splited.Length - 1];
		}

		/// <summary>
		/// 파일이름 존재하면 이름 뒤에 숫자 추가(원본 파일, jpg파일, 썸네일))
		/// </summary>
		/// <param name="doesMakeDIR">디렉토리 생성 여부</param>
		public static string GetUniqueFileName(string filePath, string filename, bool doesMakeDIR)
		{
			string FileExe;
			string FileNm;
			string FileNm_inc;

			filePath = GetValidPath(filePath);

			if (doesMakeDIR)
			{
				MakeDir(filePath);
			}

			try
			{
				FileExe = System.IO.Path.GetExtension(filename);//"."포함한 확장자
				FileNm = System.IO.Path.GetFileNameWithoutExtension(filename);//확장자 제외한 파일명
				FileNm_inc = FileNm;//번호 붙여질 파일이름
			}
			catch (Exception ex)
			{
				new TextTool.Util().writeLog(string.Format("ERR GET UNIQUE FILENAME : {0}\r\n{1}", filename, ex.ToString()));
				return filePath + System.DateTime.Now.ToString("yyMMdd_hhmmss_ms");
			}

			int nCnt = 1;

			//원본 파일, jpg파일, 썸네일(_tmb.jpg) 파일 모두 존재하지 않는 파일이름으로 생성
			while (System.IO.File.Exists(filePath + FileNm_inc + FileExe) || System.IO.File.Exists(filePath + FileNm_inc + ".jpg") || System.IO.File.Exists(filePath + FileNm_inc + "_tmb.jpg"))
			{
				FileNm_inc = string.Format("{0}{1}",
						FileNm,
						nCnt);
				nCnt++;
			}

			return filePath + FileNm_inc + FileExe;
		}

		/// <summary>
		/// 경로 뒤쪽에 디렉터리 수준 구분 문자를 검사, 추가
		/// </summary>
		public static string GetValidPath(string filePath)
		{
			filePath = filePath.Trim();
			if (filePath.Length < 1) return filePath;

			if (filePath[filePath.Length - 1] == System.IO.Path.DirectorySeparatorChar)
			{
				return filePath;
			}
			else
			{
				return filePath + System.IO.Path.DirectorySeparatorChar;
			}
		}

		/// <summary>
		/// 디렉토리 생성, 성공 true, 실패 false 리턴
		/// </summary>
		public static bool MakeDir(string Dirpath)
		{
			try
			{
				if (!System.IO.Directory.Exists(Dirpath))
				{
					System.IO.Directory.CreateDirectory(Dirpath);
				}
			}
			catch (Exception ex)
			{
				new TextTool.Util().writeLog("MakeDir ERR" + ex.ToString());
				return false;
			}
			return true;
		}

		/// <summary>
		/// 다운로드 실패시 false
		/// </summary>
		public static bool DownloadFromUrl(string urlPath, string saveFileFullName)
		{
			bool Result = true;

			using (System.Net.WebClient wc = new System.Net.WebClient())
			{
				try
				{
					wc.DownloadFile(urlPath, saveFileFullName);
				}
				catch (Exception ex)
				{
					new TextTool.Util().writeLog("ERR DOWNLOAD FILE : " + urlPath + "\r\n" + ex.ToString());
					Result = false;
				}
			}

			return Result;
		}

		/// <summary>
		/// HTTP_REFERER를 체크하는 사이트에서 사용
		/// </summary>
		public static void DownloadFromUrl(string urlPath, string referer_Path, string saveFileFullName)
		{
			using (System.Net.WebClient wc = new System.Net.WebClient())
			{
				HttpWebRequest objReq = (HttpWebRequest)WebRequest.Create(referer_Path);
				objReq.Referer = referer_Path;

				wc.Headers = objReq.Headers;

				try
				{
					wc.DownloadFile(urlPath, saveFileFullName);
				}
				catch (Exception ex)
				{
					new TextTool.Util().writeLog("ERR DOWNLOAD FILE : " + urlPath + "\r\n" + ex.ToString());
				}
			}
		}

		/// <summary>
		/// "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
		/// </summary>
		/// <param name="inputNum">UInt16~UInt64까지 가능</param>
		public static string toBase62(UInt64 inputNum)
		{
			string charSet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
			string Result = "";

			while (inputNum != 0)
			{
				Result += charSet[(int)(inputNum % 62)];
				inputNum /= 62;
			}

			return Result;
		}
		/// <summary>
		/// 로그 파일 기록
		/// </summary>
		public static void writeLog(string str)
		{
			WriteLog(new TextTool.Util().getAppCfg("logFile"), str);
		}

		/// <summary>
		/// 로그 파일 기록
		/// </summary>
		public static void writeLog(string fileName, string str)
		{
			WriteLog(fileName, str);
		}

		private static void WriteLog(string fileName, string str)
		{
			Random rdm = new Random();

			string logFile = fileName;
			string logDate = System.DateTime.Now.ToString("yyyy-MM-dd");
			string timeStamp = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

			bool wait = true;
			int loopCnt = 0;
			int sleepSec;

			if (logFile == null)
				logFile = string.Format("{0}\\log", System.Environment.CurrentDirectory);
			logFile = string.Format("{0}[{1}].txt", logFile, logDate);

			while (wait && loopCnt < 50)
			{
				try
				{
					using (StreamWriter objWriter = System.IO.File.AppendText(logFile))
					{
						objWriter.WriteLine("[{0}] {1}", timeStamp, str);
						objWriter.WriteLine();
						objWriter.Close();
					}
					wait = false;
					loopCnt++;
				}
				catch (SystemException ex)
				{
					//엑세스 이상
					wait = true;
					loopCnt++;
					sleepSec = rdm.Next(1, 21) * (10 + loopCnt) / rdm.Next(1, 10);
					System.Threading.Thread.Sleep(sleepSec);
				}
			}
		}
	}
}
