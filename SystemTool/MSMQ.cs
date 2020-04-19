using System;
using System.Messaging;
using System.Collections.Generic;
using System.Text;

namespace Tools.SystemTool
{
	public class MSMQ
	{
		/// <summary>
		/// 큐 이름 형식이 맞으면 true
		/// </summary>
		/// <param name="queuePath"></param>
		/// <returns></returns>
		private bool isQueueName(string queuePath)
		{
			MessageQueue myQ;
			bool Result;
			try
			{
				myQ = new MessageQueue(queuePath);
				Result = true;
			}
			catch
			{
				Result = false;
			}
			return Result;
		}
		/// <summary>
		/// 큐를 읽을 수 있으면 true
		/// </summary>
		public bool canReadQueue(string queuePath)
		{
			bool Result;

			if (isQueueName(queuePath))
			{
				MessageQueue myQ = new MessageQueue(queuePath);
				try
				{
					Result = myQ.CanRead;
				}
				catch
				{
					Result = false;
				}
			}
			else
				Result = false;

			return Result;
		}
		/// <summary>
		/// 해당 큐에 있는 메시지 갯수, 에러 발생시 -1
		/// </summary>
		public int queueMsgCnt(string queuePath)
		{
			int cnt = 0;

			try
			{
				MessageQueue myQ = myQ = new MessageQueue(queuePath);
				MessageEnumerator msgEnu = myQ.GetMessageEnumerator2();

				while (msgEnu.MoveNext())
					cnt++;
			}
			catch (Exception ex)
			{
				FileTool.writeLog("ERR GET QUEUE MSG COUNT : " + ex.ToString());
				cnt = -1;
			}

			return cnt;
		}
		public void clearMSMQ(string queuePath)
		{
			MessageQueue myQ = new MessageQueue(queuePath);
			int i;
			int msgCnt = queueMsgCnt(queuePath);
			for (i = 0; i < msgCnt; i++)
			{
				myQ.Receive();
			}
		}
		/// <summary>
		/// 오브젝트를 해당 큐로 전송
		/// </summary>
		public void sendData(string queuePath, object obj)
		{
			MessageQueue myQ = new MessageQueue(queuePath);
			myQ.Send(obj);
		}
		/// <summary>
		/// 해당 큐에서 설정한 오브젝트 타입으로 수신, 시간이 초과되면 null 리턴
		/// </summary>
		/// <param name="timeOut">타임아웃(초)</param>
		/// <returns></returns>
		public object receiveData(string queuePath, Type[] objectType, int timeOut)
		{
			MessageQueue myQ = new MessageQueue(queuePath);
			object rcvObj;

			myQ.Formatter = new XmlMessageFormatter(objectType);
			try
			{
				rcvObj = myQ.Receive(new TimeSpan(0, 0, timeOut)).Body;
			}
			catch
			{
				rcvObj = null;
			}

			return rcvObj;
		}
		/// <summary>
		/// 해당 큐에서 설정한 오브젝트 타입으로 큐 검색, 시간이 초과되면 null 리턴
		/// </summary>
		/// <param name="timeOut">타임아웃(초)</param>
		/// <returns></returns>
		public object peekData(string queuePath, Type[] objectType, int timeOut)
		{
			MessageQueue myQ = new MessageQueue(queuePath);
			object rcvObj;

			myQ.Formatter = new XmlMessageFormatter(objectType);
			try
			{
				rcvObj = myQ.Peek(new TimeSpan(0, 0, timeOut)).Body;
			}
			catch
			{
				rcvObj = null;
			}

			return rcvObj;
		}
		/// <summary>
		/// 해당 큐에서 설정한 오브젝트 타입으로 비동기 수신 시작
		/// </summary>
		public IAsyncResult startAsyncReceiveData(string queuePath, Type[] objectType)
		{
			MessageQueue asyncQ = new MessageQueue(queuePath);
			asyncQ.Formatter = new XmlMessageFormatter(objectType);
			return asyncQ.BeginReceive();
		}
		/// <summary>
		/// 해당 큐의 비동기 수신 종료
		/// </summary>
		public object endAsyncReceiveData(string queuePath, IAsyncResult asyncResult)
		{
			MessageQueue asyncQ = new MessageQueue(queuePath);
			return asyncQ.EndReceive(asyncResult).Body;
		}
	}
}
