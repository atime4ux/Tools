using System;
using System.Data;
using System.Data.Odbc;

namespace Tools.DbTool.Access
{
	public class OdbcAccess : IDatabase
	{
		private string ConnectionString { get; set; }

		public OdbcAccess(string conStr)
		{
			ConnectionString = conStr;
		}

		public IDbConnection CreateConnection()
		{
			return new OdbcConnection(ConnectionString);
		}

		public void CloseConnection(IDbConnection con)
		{
			OdbcConnection objCon = (OdbcConnection)con;
			objCon.Close();
			objCon.Dispose();
		}

		public IDbCommand CreateCommand(IDbConnection con, CommandType commandType, string commandText)
		{
			return new OdbcCommand { 
				Connection = (OdbcConnection)con,
				CommandText = commandText,
				CommandType = commandType
			};
		}

		public IDbDataAdapter CreateAdapter(IDbCommand command)
		{
			return new OdbcDataAdapter((OdbcCommand)command);
		}

		public IDbDataParameter CreateParam(IDbCommand command)
		{
			return ((OdbcCommand)command).CreateParameter();
		}

		public IDbDataParameter CreateParam(string name, object value)
		{
			return new OdbcParameter
			{
				ParameterName = name,
				Value = value
			};
		}

		public IDbDataParameter CreateParam(string name, object value, DbType dbType)
		{
			return new OdbcParameter
			{
				DbType = dbType,
				ParameterName = name,
				Value = value
			};
		}

		public IDbDataParameter CreateParam(string name, object value, DbType dbType, ParameterDirection direction)
		{
			return new OdbcParameter
			{
				DbType = dbType,
				ParameterName = name,
				Direction = direction,
				Value = value
			};
		}

		public IDbDataParameter CreateParam(string name, int size, object value, DbType dbType, ParameterDirection direction)
		{
			return new OdbcParameter
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