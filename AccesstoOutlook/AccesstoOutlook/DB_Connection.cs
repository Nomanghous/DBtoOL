using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Configuration;

namespace AccesstoOutlook
{
    class DB_Connection
    {

        private static OleDbConnection _connection;
        private DB_Connection()
        { 
        }
        public static OleDbConnection GetDBConnection()
        {
          
            if (_connection == null){
                _connection = new OleDbConnection(Passvalues.connectionString);
                return _connection;
            }
            else if (_connection.ConnectionString == "")
            {
                _connection.Dispose();
                _connection = new OleDbConnection(Passvalues.connectionString);
                return _connection;
            }
            else
            return _connection;

        }
    }
}
