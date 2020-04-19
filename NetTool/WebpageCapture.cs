using System;
using System.IO;
using System.Drawing;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace Tools.NetTool
{
	/// <summary>
	/// clsWeb.WebpageCapture()에서 사용
	/// </summary>
	class WebpageCapture
	{
		System.Threading.Thread t1;
		private WebBrowser wb;

		private string URL;
		private string fileDir;
		private string filename;
		private string fileExtension;

		//저장된 파일 경로
		private string filename_final = null;
		public string filepath
		{
			get
			{
				return filename_final;
			}
		}

		//이미지 형식
		private System.Drawing.Imaging.ImageFormat objImgFormat;

		//크기 비교를 위한 메모리 스트림
		private MemoryStream MS_final;
		private MemoryStream MS_1;
		private MemoryStream MS_2;

		//특정부분 캡쳐 변수
		private bool isFullScreen;
		private int x;
		private int y;
		private int width;
		private int height;

		public WebpageCapture(string URL, string filename, System.Drawing.Imaging.ImageFormat objImgFormat)
		{
			this.URL = URL;
			this.filename = filename;
			this.objImgFormat = objImgFormat;
			this.isFullScreen = true;
			this.x = 0;
			this.y = 0;
			this.width = 0;
			this.height = 0;
		}

		public WebpageCapture(string URL, string filename, int x, int y, int width, int height, System.Drawing.Imaging.ImageFormat objImgFormat)
		{
			this.URL = URL;
			this.filename = filename;
			this.objImgFormat = objImgFormat;
			this.isFullScreen = false;
			this.x = x;
			this.y = y;
			this.width = width;
			this.height = height;
		}

		public void WebpageCapture_noReturn()
		{
			t1 = new System.Threading.Thread(new System.Threading.ThreadStart(Capture));
			t1.SetApartmentState(System.Threading.ApartmentState.STA);
			t1.Start();
		}

		public string RunCapture()
		{
			Capture();
			return this.filepath;
		}

		void wb_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
		{
			//페이지 오류
			if (wb.DocumentType.Equals("파일"))
			{
				MS_final = new MemoryStream();
				FileTool.writeLog("CAPTURE FAIL : NOT EXIST PAGE");
				return;
			}

			Bitmap bmp;

			int wb_width;
			int wb_height;

			//스크롤바만큼 크기 증가
			wb_width = wb.Document.Body.ScrollRectangle.Size.Width + 20;
			wb_height = wb.Document.Body.ScrollRectangle.Size.Height + 20;
			wb.Size = new Size(wb_width, wb_height);

			wb_width = wb.Document.Body.ScrollRectangle.Size.Width;
			wb_height = wb.Document.Body.ScrollRectangle.Size.Height;

			//전체 이미지 생성
			bmp = new Bitmap(wb_width, wb_height);
			wb.DrawToBitmap(bmp, new Rectangle(0, 0, wb_width, wb_height));

			//부분 이미지 생성
			if (!isFullScreen)
			{
				Bitmap bmp2 = new Bitmap(width, height);
				Graphics objGRP = Graphics.FromImage((Image)bmp2);

				objGRP.DrawImage((Image)bmp, 0, 0, new Rectangle(x, y, width, height), GraphicsUnit.Pixel);
				objGRP.Dispose();
				bmp.Dispose();
				bmp = bmp2;
			}

			//스트림 크기 비교
			if (MS_1.Length == 0)
				bmp.Save(MS_1, objImgFormat);
			else
				bmp.Save(MS_2, objImgFormat);

			//더 큰 사이즈 이미지로 치환
			if (MS_2.Length >= MS_1.Length)
			{
				ClearMemoryStream(MS_1);
				MS_1 = new MemoryStream();
				MS_2.WriteTo(MS_1);
				ClearMemoryStream(MS_2);
				MS_2 = new MemoryStream();
			}

			//최종 이미지 저장
			if (wb.ReadyState == WebBrowserReadyState.Complete)
			{
				MS_final = new MemoryStream();
				MS_1.WriteTo(MS_final);
			}

			bmp.Dispose();
		}

		private void ClearMemoryStream(MemoryStream MS)
		{
			MS.SetLength(0);
			MS.Close();
			MS.Dispose();
			MS = null;
		}

		private void Capture()
		{
			wb = new WebBrowser();
			wb.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(wb_DocumentCompleted);

			System.Diagnostics.Stopwatch objWatch = new System.Diagnostics.Stopwatch();
			int waitTime = 5;

			MS_1 = new MemoryStream();
			MS_2 = new MemoryStream();
			MS_final = null;

			if (URL.Length > 0 && filename.Length > 0)
			{
				try
				{
					wb.ScrollBarsEnabled = false;
					wb.ScriptErrorsSuppressed = true;
					wb.Navigate(URL);

					//대기
					objWatch.Start();
					while (MS_final == null)
					{
						Application.DoEvents(); //메시지로 인해 이벤트 발생 시 함수 호출 가능
						if (objWatch.Elapsed.Seconds > waitTime)
						{
							//응답없으면 다시 호출
							wb.Stop();
							wb.Navigate(URL);
							if (waitTime < 30)
							{
								waitTime = waitTime + 5;
							}
							else
							{
								//30초동안 처리 안되면 루프 탈출
								if (MS_1.Length > 0)
								{
									MS_final = new MemoryStream();
									MS_1.WriteTo(MS_final);
									ClearMemoryStream(MS_1);
								}
								FileTool.writeLog("TIMEOUT WEBPAGE CAPTURE : " + waitTime + "sec");
								break;
							}
						}
						System.Threading.Thread.Sleep(1000);
					}
					objWatch.Stop();
					wb.Stop();

					if (MS_final != null)
						saveImg();
				}
				catch (Exception ex)
				{
					FileTool.writeLog("FAIL WEBPAGE CAPTURE : " + ex.ToString());
					filename_final = "FAIL";
				}
			}

			ClearMemoryStream(MS_1);
			ClearMemoryStream(MS_2);

			((Control)wb).Enabled = false;
			wb.Dispose();
			wb = null;
		}

		//이미지 저장
		private void saveImg()
		{
			if (MS_final.Length > 0)
			{
				int i = 0;

				fileDir = filename.Substring(0, filename.LastIndexOf("\\") + 1);
				fileExtension = System.IO.Path.GetExtension(filename);
				filename = System.IO.Path.GetFileNameWithoutExtension(filename);

				filename_final = fileDir + filename + fileExtension;

				try
				{
					while (System.IO.File.Exists(filename_final))
					{
						i++;
						filename_final = fileDir + filename + "_" + i + fileExtension;
					}
					Bitmap.FromStream(MS_final).Save(filename_final, objImgFormat);
				}
				catch (Exception ex)
				{
					FileTool.writeLog("FAIL FILE SAVE : " + ex.ToString());
					filename_final = "FAIL";
				}
			}
			else
				FileTool.writeLog("IMAGE STREAM IS EMPTY");

			ClearMemoryStream(MS_final);
		}
	}
}
