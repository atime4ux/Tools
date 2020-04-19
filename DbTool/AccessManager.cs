using System;
using System.Data;
using System.Configuration;

namespace Tools.DbTool
{
  public class AccessManager
  {
    private Access.IDatabase database;
    private string providerName;
    private bool WriteLog
    {
      get
      {
        return false;
      }
    }

    public AccessManager(string connectionName)
    {
      InitByKey(connectionName);
    }

    public AccessManager(string connectionString, string provider)
    {
      InitByConnectionString(connectionString, provider);
    }

    private void InitByConnectionString(string conStr, string provider)
    {
      providerName = provider.ToLower();
      switch (providerName)
      {
        case "sql":
          database = new Access.SqlServerAccess(conStr);
          break;
        case "oracle":
          database = new Access.OracleAccess(conStr);
          break;
        case "oleDb":
          database = new Access.OledbAccess(conStr);
          break;
        case "odbc":
          database = new Access.OdbcAccess(conStr);
          break;
      }
    }

    private void InitByKey(string key)
    {
      ConnectionStringSettings connectionInfo = ConfigurationManager.ConnectionStrings[key];
      InitByConnectionString(connectionInfo.ConnectionString, connectionInfo.ProviderName);
    }

    public IDbConnection GetDatabasecOnnection()
    {
      return database.CreateConnection();
    }

    public void CloseConnection(IDbConnection connection)
    {
      database.CloseConnection(connection);
    }

    public IDbDataParameter CreateParameter(string name, object value)
    {
      return database.CreateParam(name, value);
    }

    public IDbDataParameter CreateParameter(string name, object value, DbType dbType)
    {
      return database.CreateParam(name, value, dbType);
    }

    public IDbDataParameter CreateParameter(string name, object value, DbType dbType, ParameterDirection direction)
    {
      return database.CreateParam(name, value, dbType, direction);
    }

    public IDbDataParameter CreateParameter(string name, object value, int size, DbType dbType, ParameterDirection direction)
    {
      return database.CreateParam(name, size, value, dbType, direction);
    }

    private void SuccessLog(string query)
    {
      if (WriteLog)
      {
        new TextTool.Util().writeLog($"DB SUCCESS : {query}");
      }
    }

    private void FailLog(string failMsg, string query = "")
    {
      if (WriteLog)
      {
        string result;
        if (query.Length > 0)
        {
          result = $"DBEXEC ERR : {failMsg}\r\n{query}";
        }
        else
        {
          result = $"DBEXEC ERR : {failMsg}";
        }

        new TextTool.Util().writeLog(result);
      }
    }

    private void OpenConnection(IDbConnection con)
    {
      try
      {
        con.Open();
      }
      catch (Exception ex)
      {
        FailLog($"Connection FAIL : {ex.Message}", "");
        throw ex;
      }
    }

    public DataSet GetDataSet(string commandText, IDbDataParameter[] parameters, CommandType commandType = CommandType.Text)
    {
      using (var connection = database.CreateConnection())
      {
        OpenConnection(connection);

        using (var command = database.CreateCommand(connection, commandType, commandText))
        {
          if (parameters != null)
          {
            foreach (var parameter in parameters)
            {
              command.Parameters.Add(parameter);
            }
          }

          var dataset = new DataSet();
          var dataAdaper = database.CreateAdapter(command);

          try
          {
            dataAdaper.Fill(dataset);
            SuccessLog(commandText);
          }
          catch (Exception ex)
          {
            FailLog(commandText);
            throw ex;
          }

          return dataset;
        }
      }
    }

    public DataTable GetDataTable(string commandText, IDbDataParameter[] parameters, CommandType commandType = CommandType.Text)
    {
      return GetDataSet(commandText, parameters, commandType).Tables[0];
    }

    public IDataReader GetDataReader(out IDbConnection connection, string commandText, IDbDataParameter[] parameters, CommandType commandType)
    {
      IDataReader reader = null;

      connection = database.CreateConnection();
      OpenConnection(connection);

      var command = database.CreateCommand(connection, commandType, commandText);
      if (parameters != null)
      {
        foreach (var parameter in parameters)
        {
          command.Parameters.Add(parameter);
        }
      }

      try
      {
        reader = command.ExecuteReader();
        SuccessLog(commandText);
      }
      catch (Exception ex)
      {
        FailLog(commandText);
        throw ex;
      }

      return reader;
    }

    public void ExecuteNonQuery(string commandText, IDbDataParameter[] parameters, CommandType commandType = CommandType.Text)
    {
      using (var connection = database.CreateConnection())
      {
        OpenConnection(connection);

        using (var command = database.CreateCommand(connection, commandType, commandText))
        {
          if (parameters != null)
          {
            foreach (var parameter in parameters)
            {
              command.Parameters.Add(parameter);
            }
          }

          try
          {
            command.ExecuteNonQuery();
            SuccessLog(commandText);
          }
          catch (Exception ex)
          {
            FailLog(commandText);
            throw ex;
          }
        }
      }
    }

    public object ExecuteScalar(string commandText, IDbDataParameter[] parameters, CommandType commandType = CommandType.Text)
    {
      using (var connection = database.CreateConnection())
      {
        OpenConnection(connection);

        using (var command = database.CreateCommand(connection, commandType, commandText))
        {
          if (parameters != null)
          {
            foreach (var parameter in parameters)
            {
              command.Parameters.Add(parameter);
            }
          }

          object result;
          try
          {
            result = command.ExecuteScalar();
            SuccessLog(commandText);
          }
          catch (Exception ex)
          {
            FailLog(commandText);
            throw ex;
          }

          return result;
        }
      }
    }

    public void ExecuteNonQueryTran(string commandText, IDbDataParameter[] parameters, CommandType commandType = CommandType.Text)
    {
      IDbTransaction transactionScope;
      using (var connection = database.CreateConnection())
      {
        OpenConnection(connection);
        transactionScope = connection.BeginTransaction();

        using (var command = database.CreateCommand(connection, commandType, commandText))
        {
          if (parameters != null)
          {
            foreach (var parameter in parameters)
            {
              command.Parameters.Add(parameter);
            }
          }

          try
          {
            command.ExecuteNonQuery();
            transactionScope.Commit();
            SuccessLog(commandText);
          }
          catch (Exception)
          {
            transactionScope.Rollback();
            FailLog(commandText);
          }
          finally
          {
            connection.Close();
          }
        }
      }
    }

    public void ExecuteNonQueryTran(string commandText, IDbDataParameter[] parameters, IsolationLevel isolationLevel, CommandType commandType = CommandType.Text)
    {
      IDbTransaction transactionScope;
      using (var connection = database.CreateConnection())
      {
        OpenConnection(connection);
        transactionScope = connection.BeginTransaction(isolationLevel);

        using (var command = database.CreateCommand(connection, commandType, commandText))
        {
          if (parameters != null)
          {
            foreach (var parameter in parameters)
            {
              command.Parameters.Add(parameter);
            }
          }

          try
          {
            command.ExecuteNonQuery();
            transactionScope.Commit();
            SuccessLog(commandText);
          }
          catch (Exception)
          {
            transactionScope.Rollback();
            FailLog(commandText);
          }
          finally
          {
            connection.Close();
          }
        }
      }
    }
  }
}