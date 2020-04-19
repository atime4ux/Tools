using System;
using System.Data;
using Oracle.ManagedDataAccess.Client;

namespace Tools.DbTool.Access
{
	public class OracleAccess : IDatabase
	{
		private string ConnectionString { get; set; }

		public OracleAccess(string conStr)
		{
			ConnectionString = conStr;
		}

		public IDbConnection CreateConnection()
		{
			return new OracleConnection(ConnectionString);
		}

		public void CloseConnection(IDbConnection con)
		{
			OracleConnection objCon = (OracleConnection)con;
			objCon.Close();
			objCon.Dispose();
		}

		public IDbCommand CreateCommand(IDbConnection con, CommandType commandType, string commandText)
		{
			return new OracleCommand { 
				Connection = (OracleConnection)con,
				CommandText = commandText,
				CommandType = commandType
			};
		}

		public IDbDataAdapter CreateAdapter(IDbCommand command)
		{
			return new OracleDataAdapter((OracleCommand)command);
		}

		public IDbDataParameter CreateParam(IDbCommand command)
		{
			return ((OracleCommand)command).CreateParameter();
		}

		public IDbDataParameter CreateParam(string name, object value)
		{
			return new OracleParameter
			{
				ParameterName = name,
				Value = value
			};
		}

		public IDbDataParameter CreateParam(string name, object value, DbType dbType)
		{
			return new OracleParameter
			{
				DbType = dbType,
				ParameterName = name,
				Value = value
			};
		}

		public IDbDataParameter CreateParam(string name, object value, DbType dbType, ParameterDirection direction)
		{
			return new OracleParameter
			{
				DbType = dbType,
				ParameterName = name,
				Direction = direction,
				Value = value
			};
		}

		public IDbDataParameter CreateParam(string name, int size, object value, DbType dbType, ParameterDirection direction)
		{
			return new OracleParameter
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