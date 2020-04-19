using System.Data;

namespace Tools.DbTool.Access
{
	public interface IDatabase
	{
		IDbConnection CreateConnection();

		void CloseConnection(IDbConnection con);

		IDbCommand CreateCommand(IDbConnection con, CommandType commandType, string commandText);

		IDbDataAdapter CreateAdapter(IDbCommand command);

		IDbDataParameter CreateParam(IDbCommand command);

		IDbDataParameter CreateParam(string name, object value);

		IDbDataParameter CreateParam(string name, object value, DbType dbType);

		IDbDataParameter CreateParam(string name, object value, DbType dbType, ParameterDirection direction);

		IDbDataParameter CreateParam(string name, int size, object value, DbType dbType, ParameterDirection direction);
	}
}