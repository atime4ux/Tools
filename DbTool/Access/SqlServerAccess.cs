using System;
using System.Data;
using System.Data.SqlClient;

namespace Tools.DbTool.Access
{
	public class SqlServerAccess : IDatabase
	{
		private string ConnectionString { get; set; }

		public SqlServerAccess(string conStr)
		{
			ConnectionString = conStr;
		}

		public IDbConnection CreateConnection()
		{
			return new SqlConnection(ConnectionString);
		}

		public void CloseConnection(IDbConnection con)
		{
			SqlConnection objCon = (SqlConnection)con;
			objCon.Close();
			objCon.Dispose();
		}

		public IDbCommand CreateCommand(IDbConnection con, CommandType commandType, string commandText)
		{
			return new SqlCommand
			{
				Connection = (SqlConnection)con,
				CommandText = commandText,
				CommandType = commandType
			};
		}

		public IDbDataAdapter CreateAdapter(IDbCommand command)
		{
			return new SqlDataAdapter((SqlCommand)command);
		}

		public IDbDataParameter CreateParam(IDbCommand command)
		{
			return ((SqlCommand)command).CreateParameter();
		}

		public IDbDataParameter CreateParam(string name, object value)
		{
			return new SqlParameter
			{
				ParameterName = name,
				Value = value
			};
		}

		public IDbDataParameter CreateParam(string name, object value, DbType dbType)
		{
			return new SqlParameter
			{
				DbType = dbType,
				ParameterName = name,
				Value = value
			};
		}

		public IDbDataParameter CreateParam(string name, object value, DbType dbType, ParameterDirection direction)
		{
			return new SqlParameter
			{
				DbType = dbType,
				ParameterName = name,
				Direction = direction,
				Value = value
			};
		}

		public IDbDataParameter CreateParam(string name, int size, object value, DbType dbType, ParameterDirection direction)
		{
			return new SqlParameter
			{
				DbType = dbType,
				Size = size,
				ParameterName = name,
				Direction = direction,
				Value = value
			};
		}

		public void BulkCopy(SqlConnection connection, DataTable dt, string destinationTable)
		{
			connection.Open();

			try
			{
				var sqlBulkCopy = new SqlBulkCopy(connection);
				sqlBulkCopy.DestinationTableName = destinationTable;

				foreach (DataColumn column in dt.Columns)
				{
					sqlBulkCopy.ColumnMappings.Add(column.ColumnName, column.ColumnName);
				}

				sqlBulkCopy.WriteToServer(dt);
			}
			catch (Exception ex)
			{
				throw ex;
			}
			finally
			{
				connection.Close();
			}
		}
	}
}