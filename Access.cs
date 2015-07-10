using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using Microsoft.Win32;

namespace Free.Database
{
	// TODO: Fix syntax (' vs. ")
	// TODO: Fix syntax (always close using ';' )

	public class Access
	{
		static string keyName="Software\\Free Framework\\Free.Database.Access";

		/// <summary>
		/// Returns the name of the Microsoft Access Driver, if overridden in registry. Empty string if not.
		/// </summary>
		static string GetOverriddenMicrosoftAccessDriver()
		{
			RegistryKey kDlg=Registry.CurrentUser.OpenSubKey(keyName, true);
			if(kDlg==null) kDlg=Registry.CurrentUser.CreateSubKey(keyName);

			string key=(IntPtr.Size==8)?"DriverName64":"DriverName32";
			string drivername=(string)kDlg.GetValue(key, "");

			kDlg.Close();

			return drivername;
		}

		/// <summary>
		/// Returns true if the name of the Microsoft Access Driver is the only one to be used.
		/// </summary>
		static bool IsForceOverriddenMicrosoftAccessDriver()
		{
			RegistryKey kDlg=Registry.CurrentUser.OpenSubKey(keyName, true);
			if(kDlg==null) kDlg=Registry.CurrentUser.CreateSubKey(keyName);

			string key=(IntPtr.Size==8)?"ForceDriverName64":"ForceDriverName32";
			bool force=((int)kDlg.GetValue(key, 0)==1);

			kDlg.Close();

			return force;
		}

		/// <summary>
		/// Returns true if the name of the Microsoft Access pre-2010 Driver is also to be used.
		/// </summary>
		static bool IsForceMicrosoftAccess2000Driver()
		{
			RegistryKey kDlg=Registry.CurrentUser.OpenSubKey(keyName, true);
			if(kDlg==null) kDlg=Registry.CurrentUser.CreateSubKey(keyName);

			bool force=((int)kDlg.GetValue("Force2000DriverName", 0)==1);

			kDlg.Close();

			return force;
		}

		/// <summary>
		/// Returns true if the name of the Microsoft Access 2010 Driver is also to be used.
		/// </summary>
		static bool IsForceMicrosoftAccess2010Driver()
		{
			RegistryKey kDlg=Registry.CurrentUser.OpenSubKey(keyName, true);
			if(kDlg==null) kDlg=Registry.CurrentUser.CreateSubKey(keyName);

			bool force=((int)kDlg.GetValue("Force2010DriverName", 0)==1);

			kDlg.Close();

			return force;
		}

		/// <summary>
		/// Retrieves the Connection String
		/// </summary>
		/// <param name="filename"></param>
		/// <returns></returns>
		private static List<string> GetConnectionStrings(string filename)
		{
			string template="Driver={{{1}}};DBQ={0}";

			List<string> ret=new List<string>();

			string over=GetOverriddenMicrosoftAccessDriver();
			if(over!="") ret.Add(string.Format(template, filename, over));
			if(IsForceOverriddenMicrosoftAccessDriver()) return ret;

			List<string> drivers=FilterDriversForAccess(ListODBCDrivers());

			if(IsForceMicrosoftAccess2000Driver()&&!drivers.Contains("Microsoft Access Driver (*.mdb)"))
				drivers.Insert(0, "Microsoft Access Driver (*.mdb)");

			if(IsForceMicrosoftAccess2010Driver()&&!drivers.Contains("Microsoft Access Driver (*.mdb, *.accdb)"))
				drivers.Insert(0, "Microsoft Access Driver (*.mdb, *.accdb)");

			foreach(string driver in drivers)
				ret.Add(string.Format(template, filename, driver));

			return ret;
		}

		#region Static methods
		public static bool CreateDatabase(string filename, bool overwrite=false)
		{
			if(!overwrite&&File.Exists(filename)) return false;

			using(Stream s=typeof(Access).Assembly.GetManifestResourceStream("Free.Database.Access.Template.mdb"))
			{
				byte[] buffer=new byte[s.Length];
				s.Read(buffer, 0, (int)s.Length);
				File.WriteAllBytes(filename, buffer);
			}
			return true;
		}

		/// <summary>
		/// Converts .NET data types to SQL data types
		/// </summary>
		/// <param name="datatype"></param>
		/// <returns></returns>
		public static string GetAccessDatatype(string datatype)
		{
			switch(datatype)
			{
				case "System.Int16":
				case "System.Int32":
				case "System.Int64":
				case "System.UInt16":
				case "System.UInt32":
				case "System.UInt64": return "INTEGER";
				case "System.Decimal":
				case "System.Double": return "DOUBLE";
				case "System.String": return "TEXT";
				case "System.Byte[]": return "IMAGE";
				default: throw new Exception("Type not supported!");
			}
		}

		/// <summary>
		/// Converts .NET data types to DBType
		/// </summary>
		/// <param name="datatype"></param>
		/// <returns></returns>
		public static DbType GetDBType(string datatype)
		{
			switch(datatype)
			{
				case "System.Boolean": return DbType.Boolean;
				case "System.Int16": return DbType.Int16;
				case "System.Int32": return DbType.Int32;
				case "System.Int64": return DbType.Int64;
				case "System.UInt16": return DbType.UInt16;
				case "System.UInt32": return DbType.UInt32;
				case "System.UInt64": return DbType.UInt64;
				case "System.Decimal": return DbType.Decimal;
				case "System.Double": return DbType.Double;
				case "System.String": return DbType.String;
				case "System.Byte[]": return DbType.Binary;
				default: throw new Exception("Type not supported!");
			}
		}
		#endregion

		public bool Connected { get; private set; }
		OdbcConnection odbcConnection=null;

		#region Basics
		/// <summary>
		/// Opens an access database
		/// </summary>
		/// <param name="filename"></param>
		public bool OpenDatabase(string filename)
		{
			if(Connected) CloseDatabase(); // Close database, if still open
			Connected=false;

			List<string> connectionStrings=GetConnectionStrings(filename);

			foreach(string cs in connectionStrings)
			{
				try
				{
					odbcConnection=new OdbcConnection(cs);
					odbcConnection.Open();
					Connected=true;

					return true;
				}
				catch
				{
					Connected=false;
				}
			}

			return false;
		}

		/// <summary>
		/// Closes an access database
		/// </summary>
		public void CloseDatabase()
		{
			odbcConnection.Close();
			odbcConnection.Dispose();
			odbcConnection=null;
			Connected=false;
		}

		/// <summary>
		/// Start a transaction
		/// </summary>
		public void StartTransaction()
		{
		}

		/// <summary>
		/// Commits a transaction
		/// </summary>
		public void CommitTransaction()
		{
		}
		#endregion

		#region Data Edits
		/// <summary>
		/// Creates a table
		/// </summary>
		/// <param name="table"></param>
		/// <returns></returns>
		public void CreateTable(DataTable table)
		{
			List<string> columns=new List<string>();

			string tableName=table.TableName;
			if(tableName=="") throw new Exception("Can't create table without name!");

			for(int i=0; i<table.Columns.Count; i++)
			{
				string column=table.Columns[i].Caption;

				string type=GetAccessDatatype(table.Columns[i].DataType.FullName);
				if(type=="INTEGER"&&table.Columns[i].AutoIncrement) column+=" AUTOINCREMENT";
				else column+=" "+type;

				if(table.Columns[i].AllowDBNull==false) column+=" NOT NULL";

				columns.Add(column);
			}

			// Build SQL command
			string sqlCommand="CREATE TABLE \""+tableName+"\" (";
			for(int i=0; i<columns.Count; i++)
			{
				if(i==columns.Count-1) sqlCommand+=" "+columns[i];
				else sqlCommand+=" "+columns[i]+",";
			}

			sqlCommand+=");";

			ExecuteNonQuery(sqlCommand);
		}

		/// <summary>
		/// Deletes a table
		/// </summary>
		/// <param name="tableName"></param>
		public void RemoveTable(string tableName)
		{
			ExecuteNonQuery(string.Format("DROP TABLE \"{0}\";", tableName));
		}

		/// <summary>
		/// Deletes multiple tables
		/// </summary>
		/// <param name="tables"></param>
		public void RemoveTables(List<string> tables)
		{
			foreach(string i in tables) RemoveTable(i);
		}

		public void RemoveEntry(string table, string key, string value)
		{
			ExecuteNonQuery(string.Format("DELETE FROM \"{0}\" WHERE {1}={2};", table, key, value));
		}

		/// <summary>
		/// Inserts data to a table
		/// </summary>
		/// <param name="addData"></param>
		public void InsertData(DataTable insertData, string indexToRemove="")
		{
			string sqlCommand=string.Format("INSERT INTO \"{0}\" (", insertData.TableName);

			bool first=true;

			// Columns
			for(int i=0; i<insertData.Columns.Count; i++)
			{
				if(insertData.Columns[i].ToString()==indexToRemove) continue;

				if(first) sqlCommand+=insertData.Columns[i];
				else sqlCommand+=", "+insertData.Columns[i];

				first=false;
			}

			sqlCommand+=") VALUES";

			string tmpSqlCommand=sqlCommand;

			StartTransaction();

			// Rows (Data)
			for(int i=0; i<insertData.Rows.Count; i++)
			{
				sqlCommand=tmpSqlCommand;
				sqlCommand+=" (";

				using(OdbcCommand cmd=new OdbcCommand())
				{
					first=true;
					for(int j=0; j<insertData.Columns.Count; j++)
					{
						OdbcParameter parameter=new OdbcParameter();
						parameter.DbType=GetDBType(insertData.Columns[j].DataType.FullName);
						parameter.ParameterName=insertData.Columns[j].Caption;
						parameter.Value=insertData.Rows[i].ItemArray[j];
						if(parameter.OdbcType==OdbcType.VarBinary)
							parameter.OdbcType=OdbcType.Image; // to prevent errors for byte[511] to byte[2000]

						if(parameter.ParameterName==indexToRemove) continue;

						cmd.Parameters.Add(parameter);

						// Build SQL command
						if(first) sqlCommand+="?";
						else sqlCommand+=",?";

						first=false;
					}

					sqlCommand+=");";

					// Execute command (per Dataset)
					cmd.CommandText=sqlCommand;
					ExecuteNonQuery(cmd);
				}
			}

			CommitTransaction();
		}

		/// <summary>
		/// Updates data
		/// </summary>
		/// <param name="insertData"></param>
		public void UpdateData(DataTable updateData, string primaryKey, string tableToUpdate)
		{
			string tmpSqlCommand=string.Format("UPDATE \"{0}\" SET ", tableToUpdate);

			StartTransaction();

			// Rows (Data)
			for(int i=0; i<updateData.Rows.Count; i++)
			{
				string sqlCommand=tmpSqlCommand;

				using(OdbcCommand cmd=new OdbcCommand())
				{
					// Columns
					for(int j=0; j<updateData.Columns.Count; j++)
					{
						OdbcParameter parameter=new OdbcParameter();
						parameter.DbType=GetDBType(updateData.Columns[j].DataType.FullName);
						parameter.ParameterName=updateData.Columns[j].Caption;
						parameter.Value=updateData.Rows[i].ItemArray[j];
						cmd.Parameters.Add(parameter);

						// Build SQL command
						if(j==updateData.Columns.Count-1) sqlCommand+=string.Format("{0}=?", updateData.Columns[j].Caption);
						else sqlCommand+=string.Format("{0}=?,", updateData.Columns[j].Caption);
					}

					sqlCommand+=" WHERE "+primaryKey+" = '"+updateData.Rows[i][primaryKey].ToString()+"';";

					// Execute command (per Dataset)
					cmd.CommandText=sqlCommand;
					ExecuteNonQuery(cmd);
				}
			}

			CommitTransaction();
		}
		#endregion

		#region Queries
		/// <summary>
		/// Executes a command
		/// </summary>
		/// <param name="sqlCommand"></param>
		public int ExecuteNonQuery(string sqlCommand)
		{
			using(OdbcCommand odbcCommand=new OdbcCommand(sqlCommand, odbcConnection))
			{
				return odbcCommand.ExecuteNonQuery();
			}
		}

		/// <summary>
		/// Executes a OdbcCommand
		/// </summary>
		/// <param name="sqlCommand"></param>
		/// <returns></returns>
		public int ExecuteNonQuery(OdbcCommand sqlCommand)
		{
			sqlCommand.Connection=odbcConnection;
			return sqlCommand.ExecuteNonQuery();
		}

		/// <summary>
		/// Executes a query
		/// </summary>
		/// <param name="sqlCommand"></param>
		/// <returns></returns>
		public DataTable ExecuteQuery(string sqlCommand)
		{
			using(OdbcCommand odbcCommand=new OdbcCommand(sqlCommand, odbcConnection))
			{
				DataTable ret=new DataTable();
				OdbcDataReader tmpDataReader=odbcCommand.ExecuteReader();
				ret.Load(tmpDataReader);
				tmpDataReader.Close();
				return ret;
			}
		}

		/// <summary>
		/// Indicates whether a table exists
		/// </summary>
		/// <param name="tableName"></param>
		/// <returns></returns>
		public bool ExistsTable(string tableName)
		{
			List<string> tables=GetTables();
			if(tables.IndexOf(tableName)==-1) return false;
			return true;
		}

		/// <summary>
		/// Retrieves a list of all table names
		/// </summary>
		/// <returns></returns>
		public List<string> GetTables()
		{
			DataTable dt=odbcConnection.GetSchema("Tables");

			List<string> ret=new List<string>();

			for(int i=0; i<dt.Rows.Count; i++)
			{
				if(dt.Rows[i]["TABLE_TYPE"].ToString()=="TABLE") ret.Add(dt.Rows[i]["TABLE_NAME"].ToString());
			}

			return ret;
		}

		/// <summary>
		/// Retrieves a empty dataset for a table, e.g. for add
		/// </summary>
		/// <param name="tableName"></param>
		/// <returns></returns>
		public DataTable GetTableStructure(string tableName)
		{
			string sqlCommand="SELECT * FROM \""+tableName+"\";";
			DataTable ret=ExecuteQuery(sqlCommand);
			ret.Rows.Clear();
			ret.TableName=tableName;
			return ret;
		}

		/// <summary>
		/// Retrieves a table
		/// </summary>
		/// <param name="tblName"></param>
		/// <returns></returns>
		public DataTable GetTable(string tableName)
		{
			string sqlCommand="SELECT * FROM \""+tableName+"\";";
			DataTable ret=ExecuteQuery(sqlCommand);
			ret.TableName=tableName;
			return ret;
		}
		#endregion

#if ODBC
		[DllImport("odbc32.dll")]
		static extern int SQLDataSources(int EnvironmentHandle, int Direction,
			StringBuilder ServerName, int ServerNameBufferLenIn, ref int ServerNameBufferLenOut,
			StringBuilder Description, int DescriptionBufferLenIn, ref int DescriptionBufferLenOut);

		[DllImport("odbc32.dll")]
		static extern int SQLDrivers(int EnvironmentHandle, int Direction,
			StringBuilder DriverDescription, int DriverDescriptionBufferLenIn, ref int DriverDescriptionBufferLenOut,
			StringBuilder DriverAttributes, int DriverAttributesLenIn, ref int DriverAttributesLenOut);

		[DllImport("odbc32.dll")]
		static extern int SQLAllocEnv(ref int EnvironmentHandle);

		const int SQL_FETCH_NEXT=1;
		const int SQL_FETCH_FIRST=2;
		const int SQL_FETCH_FIRST_USER=31;
		const int SQL_FETCH_FIRST_SYSTEM=32;

		public static List<string> ListODBCSourcesAPI()
		{
			List<string> result=new List<string>();

			int environmentHandle=0;

			if(SQLAllocEnv(ref environmentHandle)!=-1)
			{
				StringBuilder serverName=new StringBuilder(1024);
				StringBuilder description=new StringBuilder(1024);
				int serverNameLen=0;
				int descriptionLen=0;

				int ret=SQLDataSources(environmentHandle, SQL_FETCH_FIRST_SYSTEM,
					serverName, serverName.Capacity, ref serverNameLen,
					description, description.Capacity, ref descriptionLen);

				while(ret==0)
				{
					result.Add(serverName+System.Environment.NewLine+description);

					ret=SQLDataSources(environmentHandle, SQL_FETCH_NEXT,
						serverName, serverName.Capacity, ref serverNameLen,
						description, description.Capacity, ref descriptionLen);
				}
			}

			return result;
		}

		public static List<string> ListODBCDriversAPI()
		{
			List<string> result=new List<string>();

			int environmentHandle=0;

			if(SQLAllocEnv(ref environmentHandle)!=-1)
			{
				StringBuilder driverDescription=new StringBuilder(1024);
				StringBuilder driverAttributes=new StringBuilder(1024);
				int driverDescriptionLen=0;
				int driverAttributesLen=0;

				int ret=SQLDrivers(environmentHandle, SQL_FETCH_FIRST,
					driverDescription, driverDescription.Capacity, ref driverDescriptionLen,
					driverAttributes, driverAttributes.Capacity, ref driverAttributesLen);

				while(ret==0)
				{
					result.Add(driverDescription+System.Environment.NewLine+driverAttributes);

					driverDescriptionLen=0;
					driverAttributesLen=0;
					ret=SQLDataSources(environmentHandle, SQL_FETCH_NEXT,
						driverDescription, driverDescription.Capacity, ref driverDescriptionLen,
						driverAttributes, driverAttributes.Capacity, ref driverAttributesLen);
				}
			}

			return result;
		}
#endif

		public static List<string> ListODBCDrivers()
		{
			List<string> ret=new List<string>();

			Microsoft.Win32.RegistryKey reg=Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"Software\ODBC\ODBCINST.INI\ODBC Drivers");
			if(reg==null) return ret;

			foreach(string name in reg.GetValueNames()) ret.Add(name);

			reg.Close();

			return ret;
		}

		public static List<string> FilterDriversForAccess(List<string> drivers)
		{
			List<string> ret=new List<string>();
			if(drivers==null) return ret;

			foreach(string driver in drivers)
			{
				string lower=driver.ToLower();
				if(lower.Contains("microsoft")&&lower.Contains("access")&&lower.Contains("*.mdb")) ret.Add(driver);
			}

			return ret;
		}
	}
}
