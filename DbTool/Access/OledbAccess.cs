using System;
using System.Data;
using System.Data.OleDb;

namespace Tools.DbTool.Access
{
	public class OledbAccess : IDatabase
	{
		private string ConnectionString { get; set; }

		public OledbAccess(string conStr)
		{
			ConnectionString = conStr;
		}

		public IDbConnection CreateConnection()
		{
			return new OleDbConnection(ConnectionString);
		}

		public void CloseConnection(IDbConnection con)
		{
			OleDbConnection objCon = (OleDbConnection)con;
			objCon.Close();
			objCon.Dispose();
		}

		public IDbCommand CreateCommand(IDbConnection con, CommandType commandType, string commandText)
		{
			return new OleDbCommand { 
				Connection = (OleDbConnection)con,
				CommandText = commandText,
				CommandType = commandType
			};
		}

		public IDbDataAdapter CreateAdapter(IDbCommand command)
		{
			return new OleDbDataAdapter((OleDbCommand)command);
		}

		public IDbDataParameter CreateParam(IDbCommand command)
		{
			return ((OleDbCommand)command).CreateParameter();
		}

		public IDbDataParameter CreateParam(string name, object value)
		{
			return new OleDbParameter
			{
				ParameterName = name,
				Value = value
			};
		}

		public IDbDataParameter CreateParam(string name, object value, DbType dbType)
		{
			return new OleDbParameter
			{
				DbType = dbType,
				ParameterName = name,
				Value = value
			};
		}

		public IDbDataParameter CreateParam(string name, object value, DbType dbType, ParameterDirection direction)
		{
			return new OleDbParameter
			{
				DbType = dbType,
				ParameterName = name,
				Direction = direction,
				Value = value
			};
		}

		public IDbDataParameter CreateParam(string name, int size, object value, DbType dbType, ParameterDirection direction)
		{
			return new OleDbParameter
			{
				DbType = dbType,
				Size = size,
				ParameterName = name,
				Direction = direction,
				Value = value
			};
		}
	}
}