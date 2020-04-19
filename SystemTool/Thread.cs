namespace Tools.SystemTool
{
	public class Thread
	{
		private delegate void SetLabelCallBack(System.Windows.Forms.Label label, string str);
		private delegate void SetTextBoxCallBack(System.Windows.Forms.TextBox Textbox, string str);
		private delegate void SetButtonCallBack(System.Windows.Forms.Button button, string str, bool On_Off);

		/// <summary>
		/// 스레드에서 라벨 텍스트 변경
		/// </summary>
		public static void SetLabel(System.Windows.Forms.Label label, string str)
		{
			if (label.InvokeRequired)
			{
				SetLabelCallBack dele = new SetLabelCallBack(SetLabel);
				label.Invoke(dele, label, str);
			}
			else
				label.Text = str;
		}
		/// <summary>
		/// 스레드에서 라벨 텍스트 변경
		/// </summary>
		public static void SetTextBox(System.Windows.Forms.TextBox Textbox, string str)
		{
			if (Textbox.InvokeRequired)
			{
				SetTextBoxCallBack dele = new SetTextBoxCallBack(SetTextBox);
				Textbox.Invoke(dele, Textbox, str);
			}
			else
				Textbox.AppendText(str);
		}
		/// <summary>
		/// 스레드에서 버튼 속성 변경
		/// </summary>
		/// <param name="str">텍스트</param>
		/// <param name="On_Off">활성/비활성</param>
		public static void buttonToggle(System.Windows.Forms.Button button, string str, bool On_Off)
		{
			if (button.InvokeRequired)
			{
				SetButtonCallBack dele = new SetButtonCallBack(buttonToggle);
				button.Invoke(dele, button, str, On_Off);
			}
			else
			{
				button.Enabled = On_Off;
				if (str.Length > 0)
					button.Text = str;
			}
		}
	}
}
