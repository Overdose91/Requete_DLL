using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Data;
using System.Security.Cryptography;
using System.IO;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel ;
using System.Drawing;


namespace Ov3r
{
    struct InfoConn
    {
        public string serverName;
        public string database;
        public string userName;
        public string password;
        public int timeOut;

        public InfoConn(string server, string dataBase, string user, string passWord, int timeOut)
        {
            this.serverName = server;
            this.database = dataBase;
            this.userName = user;
            this.password = passWord;
            this.timeOut = timeOut;
        }
    }
    public class MSSQL
    {
        private InfoConn _info = new InfoConn();
        private SqlConnection _cnSql = new SqlConnection();

        public MSSQL(string server, string dataBase, string user, string passWord, int timeOut = 15)
        {
            initializingConnection(server, dataBase, user, passWord, timeOut);

        }
        public MSSQL(SqlConnection _cnSqlExt)
        {
            this._cnSql = _cnSqlExt;            
        }

        private void initializingConnection(string server, string dataBase, string user, string passWord, int timeOut)
        {
            this._info.serverName = server;
            this._info.database = dataBase;
            this._info.userName = user;
            this._info.password = passWord;
            this._cnSql.ConnectionString = "Data Source=" + this._info.serverName + ";Initial Catalog=" + this._info.database + ";User Id=" + this._info.userName + ";Password=" + this._info.password + ";Connection Timeout = " + timeOut;
        }

        public string tryConnection()
        {
            try
            {
                _cnSql.Open();
                _cnSql.Close();
                return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                _cnSql.Close();
            }
        }
        /// <summary>
        /// Retourne toutes les tables de la base de données
        /// </summary>
        /// <param name="filter">Les table(s) à exclure </param>
        /// <returns></returns>
        public List<string> GetTables(List<string> filter = null)
        {
            List<string> listTables = new List<string>();

            SqlDataAdapter data = new SqlDataAdapter("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE Table_Type='BASE TABLE' ORDER BY TABLE_NAME", this._cnSql);
            DataTable dataColumns = new DataTable();
            data.Fill(dataColumns);

            foreach (DataRow read in dataColumns.Rows)
            {
                if (filter == null || !filter.Contains(read["TABLE_NAME"].ToString()))
                {
                    listTables.Add(read["TABLE_NAME"].ToString().ToUpper());
                }
            }
            data.Dispose();
            dataColumns.Dispose();
            return listTables;
        }
        /// <summary>
        /// Retourne tous les champs d'une table ainsi que leur type 
        /// </summary>
        /// <param name="tableName">Nom de la table</param>
        /// <param name="filter">Champ(s) à exclure</param>
        /// <returns></returns>
        public IDictionary<string, string> GetColumnsType(string tableName, List<string> filter = null)
        {
            IDictionary<string, string> listFieldsTypes = new Dictionary<string, string>();
            SqlDataAdapter data = new SqlDataAdapter("SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE Table_Name='" + tableName.ToUpper() + "'", this._cnSql);
            DataTable dataColumns = new DataTable();
            data.Fill(dataColumns);

            foreach (DataRow read in dataColumns.Rows)
            {
                if (filter == null || !filter.Contains(read["COLUMN_NAME"].ToString()))
                {
                    listFieldsTypes.Add(read["COLUMN_NAME"].ToString(), read["DATA_TYPE"].ToString());
                }
            }
            data.Dispose();
            dataColumns.Dispose();
            return listFieldsTypes;
        }
        /// <summary>
        /// Retourne tous les champs d'une table
        /// </summary>
        /// <param name="tableName">Nom de la table</param>
        /// <param name="filter">Champ(s) à exclure</param>
        /// <returns></returns>
        public List<string> GetColumns(string tableName, List<string> filter = null)
        {
            List<string> listFields = new List<string>();
            SqlDataAdapter data = new SqlDataAdapter("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE Table_Name='" + tableName.ToUpper() + "'", this._cnSql);
            DataTable dataColumns = new DataTable();
            data.Fill(dataColumns);

            foreach (DataRow read in dataColumns.Rows)
            {
                if (filter == null || !filter.Contains(read["COLUMN_NAME"].ToString()))
                {
                    listFields.Add(read["COLUMN_NAME"].ToString().ToUpper());
                }
            }
            data.Dispose();
            dataColumns.Dispose();
            return listFields;
        }
        public int GetNbColumns(string tableName, List<string> filter = null)
        {
            int maxNb = 0;
            SqlDataAdapter data = new SqlDataAdapter("SELECT count(*) AS NBCOLUMN FROM INFORMATION_SCHEMA.COLUMNS WHERE Table_Name='" + tableName.ToUpper() + "'", this._cnSql);
            DataTable dataColumns = new DataTable();
            data.Fill(dataColumns);

            int.TryParse(dataColumns.Rows[0]["NBCOLUMN"].ToString(), out maxNb);
            return maxNb;
        }
        /// <summary>
        /// Retourne le nombre de table dans la base
        /// </summary>
        /// <returns></returns>
        public int GetNbTables()
        {
            int maxNb = 0;
            SqlDataAdapter data = new SqlDataAdapter("SELECT count(*) AS NBTABLE FROM INFORMATION_SCHEMA.TABLES WHERE Table_Type='BASE TABLE'", this._cnSql);
            DataTable dataIdx = new DataTable();
            data.Fill(dataIdx);

            int.TryParse(dataIdx.Rows[0]["NBTABLE"].ToString(), out maxNb);
            return maxNb;
        }
        /// <summary>
        /// Retourne la valeur max de la colonne spécifiée
        /// </summary>
        /// <param name="tableName">Nom de la table</param>
        /// <param name="columnName">Nom de la colonne</param>
        /// <returns></returns>
        public int GetValueMax(string tableName, string columnName)
        {
            int maxNb = 0;
            SqlDataAdapter data = new SqlDataAdapter("SELECT MAX(" + columnName + ") AS MAXI from " + tableName, this._cnSql);
            DataTable dataIdx = new DataTable();
            data.Fill(dataIdx);

            int.TryParse(dataIdx.Rows[0]["MAXI"].ToString(), out maxNb);
            return maxNb;
        }
        /// <summary>
        /// Retourne la valeur min de la colonne spécifiée
        /// </summary>
        /// <param name="tableName">Nom de la table</param>
        /// <param name="columnName">Nom de la colonne</param>
        /// <returns></returns>
        public int GetValueMin(string tableName, string columnName)
        {
            int minNb = 0;
            SqlDataAdapter data = new SqlDataAdapter("SELECT MIN(" + columnName + ") AS MINI FROM " + tableName, this._cnSql);
            DataTable dataIdx = new DataTable();
            data.Fill(dataIdx);

            int.TryParse(dataIdx.Rows[0]["MINI"].ToString(), out minNb);
            return minNb;
        }
        /// <summary>
        /// Retourne le nombre de ligne dans la table
        /// </summary>
        /// <param name="tableName">Nom de la table</param>
        /// <returns></returns>
        public int GetTotalRows(string tableName)
        {
            int totalRows = 0;
            SqlDataAdapter data = new SqlDataAdapter("SELECT count(*) AS totalrows FROM " + tableName, this._cnSql);
            DataTable allRows = new DataTable();
            data.Fill(allRows);

            int.TryParse(allRows.Rows[0]["TOTALROWS"].ToString(), out totalRows);
            return totalRows;
        }
        /// <summary>
        /// Supprime une ligne en fonction de l'id donné
        /// </summary>
        /// <param name="tableName">Nom de la table</param>
        /// <param name="columnName">Nom de la colonne</param>
        /// <param name="id">Numéro de l'id</param>
        public void DeleteRow(string tableName, string columnName, int id)
        {
            this._cnSql.Open();
            SqlCommand deletedRow = new SqlCommand("DELETE FROM " + tableName + " WHERE " + columnName + " = " + id, this._cnSql);
            deletedRow.ExecuteNonQuery();
            this._cnSql.Close();
        }
        /// <summary>
        /// Supprime une ligne en fonction de la chaine donné
        /// </summary>
        /// <param name="tableName">Nom de la table</param>
        /// <param name="columnName">Nom de la colonne</param>
        /// <param name="id">nom de la chaine</param>
        public void DeleteRow(string tableName, string columnName, string id)
        {
            this._cnSql.Open();
            SqlCommand deletedRow = new SqlCommand("DELETE FROM " + tableName + " WHERE " + columnName + " = '" + id + "'", this._cnSql);
            deletedRow.ExecuteNonQuery();
            this._cnSql.Close();
        }
        /// <summary>
        /// Retourne dans une List toutes les valeurs d'un colonne
        /// </summary>
        /// <param name="tableName">Nom de la table</param>
        /// <param name="column">Nom de la colonne</param>
        /// <returns></returns>
        public List<string> GetDataRow(string tableName, string column)
        {
            List<string> listRowsField = new List<string>();
            SqlDataAdapter data = new SqlDataAdapter("SELECT " + column + " as ROWDATA FROM " + tableName, this._cnSql);
            DataTable dataColumns = new DataTable();
            data.Fill(dataColumns);

            foreach (DataRow read in dataColumns.Rows)
            {
                listRowsField.Add(read["ROWDATA"].ToString());
            }
            return listRowsField;
        }
        /// <summary>
        /// Créer une nouvelle table et définit la colonne de la clé primaire
        /// </summary>
        /// <param name="tableName">Nom de la table</param>
        /// <param name="columnPrimaryKey">Nom de la colonne de la clé primaire</param>
        /// <param name="identity">Valeur de début de la clé primaire</param>
        /// <returns></returns>
        public string CreateTable(string tableName, string columnPrimaryKey, int identity = 1)
        {
            try
            {
                this._cnSql.Open();
                string createTable = "IF NOT EXISTS (select * from dbo.sysobjects where id = object_id(N'[dbo].[" + tableName + "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) ";
                createTable += "CREATE TABLE dbo." + tableName + " ([" + columnPrimaryKey + "] int IDENTITY (" + identity + ", 1) NOT NULL PRIMARY KEY); ";           
                SqlCommand req = new SqlCommand(createTable, this._cnSql);
                req.ExecuteNonQuery();
                this._cnSql.Close();
                return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                this._cnSql.Close();
            }
        }
        /// <summary>
        /// Supprimer une colonne de la table
        /// </summary>
        /// <param name="tableName">Nom de la table où sera supprimé le champ</param>
        /// <param name="columnName">Nom du champ à supprimer</param>
        public string DeleteColumn(string tableName, string columnName)
        {
            try
            {
                this._cnSql.Open();
                string deleteColumn = "IF EXISTS( SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" + tableName.ToUpper() + "' AND  COLUMN_NAME = '" + columnName.ToUpper() + "') ";
                deleteColumn += "ALTER TABLE dbo." + tableName.ToUpper() + " DROP COLUMN " + columnName.ToUpper() + "  ";
                SqlCommand req = new SqlCommand(deleteColumn, this._cnSql);
                req.ExecuteNonQuery();
                this._cnSql.Close();
                return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                this._cnSql.Close();
            }
        }
        /// <summary>
        /// Ajouter une colonne dans la table
        /// </summary>
        /// <param name="tableName">Nom de la table où sera ajouté le champ</param>
        /// <param name="columnName">Nom du champ à ajouter</param>
        /// <param name="type">Nom du type du champ (Nvarchar, int, bit etc.)</param>
        public string AddColumn(string tableName, string columnName,SqlDbType type , string isnull = "")
        {
            try
            {
                string dataType = (string.IsNullOrWhiteSpace(isnull))? "" : " ( " +isnull  + " ) "  ; 
                this._cnSql.Open();
                string addColumn = "IF NOT EXISTS( SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" + tableName.ToUpper() + "' AND  COLUMN_NAME = '" + columnName.ToUpper() + "') ";
                addColumn += "ALTER TABLE dbo." + tableName.ToUpper() + " ADD " + columnName.ToUpper() + " " + type + dataType;
                SqlCommand req = new SqlCommand(addColumn, this._cnSql);
                req.ExecuteNonQuery();
                this._cnSql.Close();
                return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                this._cnSql.Close();
            }

        }
        /// <summary>
        /// Supprime une base de données
        /// </summary>
        /// <param name="database">Nom de la base à supprimer</param>
        /// <param name="_cnSql">Connexion au sgbd</param>
        /// <returns></returns>
        public string DropDatabase(string database)
        {
            string dropBdd = " IF  EXISTS (SELECT name FROM sys.databases WHERE name = N'" + database + "') ";
            dropBdd += " DROP DATABASE [" + database + "] ";

            try
            {
                this._cnSql.Open();
                SqlCommand req = new SqlCommand(dropBdd, this._cnSql);
                req.ExecuteNonQuery();
                this._cnSql.Close();
                return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                this._cnSql.Close();
            }
        }
        /// <summary>
        /// Vérifie si le nom d'utilisateur existe déjà. 
        /// </summary>
        /// <param name="userName">Nom de l'utisateur à tester</param>
        /// <returns>Return une chaine vide s'il n'existe pas sinon 0</returns>
        public string CheckUserExist(string userName)
        {
            try
            {
                SqlCommand requete = new SqlCommand("SELECT name FROM Sys.sql_logins", this._cnSql);
                this._cnSql.Open();
                SqlDataReader read = requete.ExecuteReader();
                while (read.Read())
                {
                    if (userName.Equals(read["name"].ToString(), StringComparison.CurrentCultureIgnoreCase))
                    {
                        this._cnSql.Close();
                        return "0";
                    }
                }
                read.Close();
                this._cnSql.Close();
                return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                this._cnSql.Close();
            }
        }
        /// <summary>
        /// Créer un nouveau utilisateur
        /// </summary>
        /// <param name="user">Nom de l'utilisateur</param>
        /// <param name="password">PassWord de l'utilisateur</param>
        /// <param name="database">Nom de la base</param>
        /// <returns></returns>
        public string CreateUser(string user, string password, string database)
        {
            try
            {
                this._cnSql.Open();
                string createUSER = @"USE " + database + ";";
                createUSER += @"CREATE LOGIN " + user + " WITH PASSWORD = '" + password + "', DEFAULT_DATABASE =" + database + @", CHECK_POLICY = OFF, CHECK_EXPIRATION=OFF;";
                createUSER += @"CREATE USER [" + user + "] FOR LOGIN [" + user + "];";
                createUSER += @"EXEC sp_addrolemember N'db_owner ', N'" + user + "';";

                SqlCommand requete2 = new SqlCommand(createUSER, this._cnSql);
                requete2.ExecuteNonQuery();

                // return ("L'utilisateur a bien été ajouté !", "Utiliasteur crée...", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                this._cnSql.Close();
                return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                this._cnSql.Close();
            }

        }
        /// <summary>
        /// Vérifie si le nom de la base existe déjà. 
        /// </summary>
        /// <param name="database">Nom de la base à tester</param>
        /// <returns>Return une chaine vide s'il n'existe pas sinon 0</returns>
        public string NameBddExist(string database)
        {
            try
            {
                SqlCommand requete = new SqlCommand("SELECT name FROM sys.databases WHERE database_id > 4", this._cnSql);
                this._cnSql.Open();
                SqlDataReader read = requete.ExecuteReader();
                while (read.Read())
                {
                    if (database.Equals(read["name"].ToString(), StringComparison.CurrentCultureIgnoreCase))
                    {
                        this._cnSql.Close();
                        return "";
                    }
                }
                read.Close();
                this._cnSql.Close();
                return "0";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                this._cnSql.Close();
            }
        }
        /// <summary>
        /// Ajoute un fichier dans la base
        /// </summary>
        /// <param name="tableName">Nom de la table</param>
        /// <param name="columnName">Nom de la colonne</param>
        /// <param name="pathFile">Emplacement du fichier</param>
        /// <returns></returns>
        public string AddFile(string tableName, string columnName, string pathFile)
        {
            try 
	        {	        
		        string photoFilePath = pathFile;

                //inserer fichier dans bdd
                byte[] file = GetPhoto(photoFilePath);

                SqlCommand command = new SqlCommand("INSERT INTO " + tableName + " (" + columnName + ") Values( @file )", this._cnSql);

                command.Parameters.Add("@file", SqlDbType.Image, file.Length).Value = file;

                this._cnSql.Open();
                command.ExecuteNonQuery();
                this._cnSql.Close();
                return "";
	        }
	        catch (Exception ex)
	        {
                return ex.Message;
	        }
            finally
            {
                this._cnSql.Close();
            }
        }
        private byte[] GetPhoto(string filePath)
        {
            FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            BinaryReader reader = new BinaryReader(stream);

            byte[] photo = reader.ReadBytes((int)stream.Length);

            reader.Close();
            stream.Close();

            return photo;
        }
        /// <summary>
        /// Obtient le fichier spécifié dans la base
        /// </summary>
        /// <param name="pathNewNameFile">Emplacement de sauvegarde du fichier</param>
        /// <param name="tableName">Nom de la table</param>
        /// <param name="columnName">Nom de la colonne</param>
        /// <param name="columnWhere">La table de la condition</param>
        /// <param name="columnCondition">Le mot clé de la condition</param>
        /// <param name="buffer">Initialisation du buffer pour la taille du fichier</param>
        /// <returns></returns>
        public string GetFile(string pathNewNameFile, string tableName, string columnName, string columnWhere, string columnCondition, int buffer = 100000)
        {
            try
            {
                //recupérer fichier dans BBD
                SqlCommand selectFile = new SqlCommand("SELECT " + columnName + " AS fileDate FROM " + tableName + " WHERE " + columnWhere + " = '" + columnCondition + "'", this._cnSql);
                this._cnSql.Open();
                SqlDataReader read = selectFile.ExecuteReader();
                while (read.Read())
                {
                    string name_file = read["fileDate"].ToString();
                    int bufferSize = buffer;                   // Size of the BLOB buffer.
                    byte[] outbyte = new byte[bufferSize];

                    FileStream fs = new FileStream(pathNewNameFile, FileMode.OpenOrCreate, FileAccess.Write);
                    BinaryWriter bw = new BinaryWriter(fs);
                    long starIndex = 0;

                    // Read the bytes into outbyte[] and retain the number of bytes returned.
                    long retval = read.GetBytes(0, starIndex, outbyte, 0, bufferSize);

                    // Continue reading and writing while there are bytes beyond the size of the buffer.
                    while (retval == bufferSize)
                    {
                        bw.Write(outbyte);
                        bw.Flush();

                        // Reposition the start index to the end of the last buffer and fill the buffer.
                        starIndex += bufferSize;
                        retval = read.GetBytes(0, starIndex, outbyte, 0, bufferSize);
                    }

                    // Write the remaining buffer.
                    bw.Write(outbyte, 0, (int)retval);
                    bw.Flush();

                    // Close the output file.
                    bw.Close();
                    fs.Close();

                }

                this._cnSql.Close();
                return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                this._cnSql.Close();
            }
        }
        /// <summary>
        /// Créer une base de donnée. 
        /// </summary>
        /// <returns>Return chaine vide si la base a bien été créée.</returns>
        public string CreateBdd(string nameDatabase)
        {
            try
            {
                this._cnSql.Open();
                SqlCommand requete = new SqlCommand(@"USE master;  CREATE DATABASE " + nameDatabase, this._cnSql);
                requete.ExecuteNonQuery();

                this._cnSql.Close();
                return "";
            }
            catch (Exception ex)
            {
                return ex.Message; ;
            }
            finally
            {
                this._cnSql.Close();
            }
        }


        public SqlConnection GetConnection()
        {
            return this._cnSql;
        }
        public string Get_infoServerName()
        {
            return this._info.serverName;
        }
        public string Get_infoDatabase()
        {
            return this._info.database;
        }
        public string Get_infoUserName()
        {
            return this._info.userName;
        }
        public string Get_infoPassword()
        {
            return this._info.password;
        }
    }
    public class OLEDB
    {
        private OleDbConnection _cnOledb = new OleDbConnection();

        public OLEDB(string pathDatabase)
        {
            this._cnOledb.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data source=" + pathDatabase;
           
        }
        public OLEDB(OleDbConnection _cnOledbExt)
        {
            this._cnOledb = _cnOledbExt;
        }

        public string tryConnection()
        {
            try
            {
                this._cnOledb.Open();
                this._cnOledb.Close();
                return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                this._cnOledb.Close();
            }
        }
        /// <summary>
        /// Retourne toutes les tables de la base de données
        /// </summary>
        /// <param name="filter">Les table(s) à exclure </param>
        /// <returns></returns>
        public List<string> GetTable(List<string> filter = null)
        {
            List<string> listTables = new List<string>();
            this._cnOledb.Open();
            DataTable schemaTable = this._cnOledb.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            this._cnOledb.Close();
            foreach (DataRow TableName in schemaTable.Rows)
            {
                if (filter == null || !filter.Contains(TableName["TABLE_NAME"].ToString()))
                {
                    listTables.Add(TableName["TABLE_NAME"].ToString());
                }
            }

            return listTables;
        }
        /// <summary>
        /// Retourne toutes les colonnes d'une table
        /// </summary>
        /// <param name="tableName">Nom de la table</param>
        /// <param name="filter">Champ(s) à exclure</param>
        /// <returns></returns>
        public List<string> GetColumns(string tableName, List<string> filter = null)
        {
            List<string> listFields = new List<string>();
            this._cnOledb.Open();
            DataTable dt = this._cnOledb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, tableName,null }); 
            this._cnOledb.Close();
            foreach (DataRow TableName in dt.Rows)
            {
                if (filter == null || !filter.Contains(TableName["COLUMN_NAME"].ToString()))
                {
                    listFields.Add(TableName["COLUMN_NAME"].ToString());
                }
            }
            return listFields;
        }
        /// <summary>
        /// Retourne le nombre de colonne dans la table
        /// </summary>
        /// <returns></returns>
        public int GetNbColumns(string tableName)
        {
            int maxNb = 0;
            this._cnOledb.Open();
            DataTable dt = this._cnOledb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, tableName , null }); 
            this._cnOledb.Close();
            foreach (DataRow TableName in dt.Rows)
            {
                maxNb++;
            }
            return maxNb;
        }
        /// <summary>
        /// Retourne le nombre de table dans la base
        /// </summary>
        /// <returns></returns>
        public int GetNbTables(List<string> filter = null)
        {
            int maxNb = 0;
            this._cnOledb.Open();
            DataTable schemaTable = this._cnOledb.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            this._cnOledb.Close();
            foreach (DataRow TableName in schemaTable.Rows)
            {
                maxNb++;
            }

            return maxNb;
        }        
        /// <summary>
        /// Retourne la valeur maximale de la colonne spécifiée
        /// </summary>
        /// <param name="tableName">Nom de la table</param>
        /// <param name="columnName">Nom de la colonne</param>
        /// <returns></returns>
        public int GetValueMax(string tableName, string columnName)
        {
            int maxNb = 0;
            OleDbDataAdapter data = new OleDbDataAdapter("SELECT Max(" + columnName + ") AS maxi FROM " + tableName + ";", this._cnOledb);
            DataTable dataIdx = new DataTable();
            data.Fill(dataIdx);

            int.TryParse(dataIdx.Rows[0]["maxi"].ToString(), out maxNb);
            return maxNb;
        }
        /// <summary>
        /// Retourne la valeur minimale de la colonne spécifiée
        /// </summary>
        /// <param name="tableName">Nom de la table</param>
        /// <param name="columnName">Nom de la colonne</param>
        /// <returns></returns>
        public int GetValueMin(string tableName, string columnName)
        {
            int minNb = 0;
            OleDbDataAdapter data = new OleDbDataAdapter("SELECT MIN(" + columnName + ") AS mini FROM " + tableName, this._cnOledb);
            DataTable dataIdx = new DataTable();
            data.Fill(dataIdx);

            int.TryParse(dataIdx.Rows[0]["mini"].ToString(), out minNb);
            return minNb;
        }
        /// <summary>
        /// Retourne le nombre de ligne dans la table
        /// </summary>
        /// <param name="tableName">Nom de la table</param>
        /// <returns></returns>
        public int GetTotalRows(string tableName)
        {
            int totalRows = 0;
            OleDbDataAdapter data = new OleDbDataAdapter("SELECT count(*) AS totalrows FROM " + tableName, this._cnOledb);
            DataTable allRows = new DataTable();
            data.Fill(allRows);

            int.TryParse(allRows.Rows[0]["TOTALROWS"].ToString(), out totalRows);
            return totalRows;
        }
        /// <summary>
        /// Supprime une ligne en fonction de l'id donné
        /// </summary>
        /// <param name="tableName">Nom de la table</param>
        /// <param name="columnName">Nom de la colonne</param>
        /// <param name="id">Numéro de l'id</param>
        public void DeleteRow(string tableName, string columnName, int id)
        {
            this._cnOledb.Open();
            OleDbCommand deletedRow = new OleDbCommand("DELETE FROM " + tableName + " WHERE " + columnName + " = " + id.ToString(), this._cnOledb);
            deletedRow.ExecuteNonQuery();
            this._cnOledb.Close();
        }
        /// <summary>
        /// Supprime une ligne en fonction de la chaine donné
        /// </summary>
        /// <param name="tableName">Nom de la table</param>
        /// <param name="columnName">Nom de la colonne</param>
        /// <param name="id">Nom de la chaine</param>
        public void DeleteRow(string tableName, string columnName, string id)
        {
            this._cnOledb.Open();
            OleDbCommand deletedRow = new OleDbCommand("DELETE FROM " + tableName + " WHERE " + columnName + " = '" + id.ToString() + "'", this._cnOledb);
            deletedRow.ExecuteNonQuery();
            this._cnOledb.Close();
        }
        /// <summary>
        /// Retourne dans une List toutes les valeurs d'un colonne
        /// </summary>
        /// <param name="tableName">Nom de la table</param>
        /// <param name="column">Nom de la colonne</param>
        /// <returns></returns>
        public List<string> GetDataRow(string tableName, string column)
        {
            List<string> listRowsField = new List<string>();
            OleDbDataAdapter data = new OleDbDataAdapter("SELECT " + column + " as ROWDATA FROM " + tableName, this._cnOledb);
            DataTable dataColumns = new DataTable();
            data.Fill(dataColumns);

            foreach (DataRow read in dataColumns.Rows)
            {
                listRowsField.Add(read["ROWDATA"].ToString());
            }
            return listRowsField;
        }
        /// <summary>
        /// Créer une nouvelle table
        /// </summary>
        /// <param name="tableName">Nom de la nouvelle table</param>
        /// <returns></returns>
        public string CreateTable(string tableName, string columnName)
        {
            try
            {
                this._cnOledb.Open();
                string createTable = "CREATE TABLE " + tableName + " (" + columnName + " AUTOINCREMENT PRIMARY KEY)";
                OleDbCommand req = new OleDbCommand(createTable, this._cnOledb);
                req.ExecuteNonQuery();
                this._cnOledb.Close();
                return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                this._cnOledb.Close();
            }
        }
        /// <summary>
        /// Supprimer une colonne de la table
        /// </summary>
        /// <param name="tableName">Nom de la table où sera supprimé le champ</param>
        /// <param name="columnName">Nom du champ à supprimer</param>
        public string DeleteColumn(string tableName, string columnName)
        {
            try
            {                
                this._cnOledb.Open();
                string deleteColumn = "ALTER TABLE [" + tableName + "] DROP COLUMN [" + columnName + "];";
                OleDbCommand req = new OleDbCommand(deleteColumn, this._cnOledb);
                req.ExecuteNonQuery();
                this._cnOledb.Close();
                return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                this._cnOledb.Close();
            }
        }
        /// <summary>
        /// Ajouter une colonne dans table
        /// </summary>
        /// <param name="tableName">Nom de la table où sera supprimé le champ</param>
        /// <param name="columnName">Nom du champ à supprimer</param>
        /// <param name="type">Nom du type de la colonne</param>
        /// <param name="isnull">Valeur du type</param>/// 
        public string AddColumn(string tableName, string columnName, string type, string isnull = "")
        {
            try
            {
                string typeColumn = (string.IsNullOrWhiteSpace(isnull)) ? type : type + "(" + isnull + ")";
      

                this._cnOledb.Open();
                string addColumn = "ALTER TABLE [" + tableName + "] ADD COLUMN [" + columnName + "] " + typeColumn ;
                OleDbCommand req = new OleDbCommand(addColumn, this._cnOledb);
                req.ExecuteNonQuery();
                this._cnOledb.Close();
                return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                this._cnOledb.Close();
            }
        }
        /// <summary>
        /// Ajoute une clé primaire
        /// </summary>
        /// <param name="tableName"></param>
        /// <returns></returns>
        public string AddPrimaryKey(string tableName, string columnName)
        {
            try
            {
                this._cnOledb.Open();
                OleDbCommand cmd = new OleDbCommand("ALTER TABLE " + tableName + " ADD CONSTRAINT PK_" + tableName + " Primary Key(" + columnName + ");", this._cnOledb);
                cmd.ExecuteNonQuery();

                this._cnOledb.Close();
                return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                this._cnOledb.Close();
            }
        }
        public OleDbConnection GetConnection()
        {
            return this._cnOledb;
        }
    }
}
namespace Ov3rCrypto
{
    static class Cryptage
    {
        public static string Encrypt(string toEncrypt, string keyCryptage, bool useHashing)
        {
            byte[] keyArray;
            byte[] toEncryptArray = UTF8Encoding.UTF8.GetBytes(toEncrypt);

            string key = keyCryptage;

            //System.Windows.Forms.return (key);
            //If hashing use get hashcode regards to your key
            if (useHashing)
            {
                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));
                //Always release the resources and flush data
                // of the Cryptographic service provide. Best Practice

                hashmd5.Clear();
            }
            else
                keyArray = UTF8Encoding.UTF8.GetBytes(key);

            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
            //set the secret key for the tripleDES algorithm
            tdes.Key = keyArray;
            //mode of operation. there are other 4 modes.
            //We choose ECB(Electronic code Book)
            tdes.Mode = CipherMode.ECB;
            //padding mode(if any extra byte added)

            tdes.Padding = PaddingMode.PKCS7;

            ICryptoTransform cTransform = tdes.CreateEncryptor();
            //transform the specified region of bytes array to resultArray
            byte[] resultArray =
              cTransform.TransformFinalBlock(toEncryptArray, 0,
              toEncryptArray.Length);
            //Release resources held by TripleDes Encryptor
            tdes.Clear();
            //Return the encrypted data into unreadable string format
            return Convert.ToBase64String(resultArray, 0, resultArray.Length);
        }
        public static string Decrypt(string cipherString, string keyCryptage, bool useHashing)
        {
            byte[] keyArray;
            //get the byte code of the string

            byte[] toEncryptArray = Convert.FromBase64String(cipherString);

            string key = keyCryptage;

            if (useHashing)
            {
                //if hashing was used get the hash code with regards to your key
                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));
                //release any resource held by the MD5CryptoServiceProvider

                hashmd5.Clear();
            }
            else
            {
                //if hashing was not implemented get the byte code of the key
                keyArray = UTF8Encoding.UTF8.GetBytes(key);
            }

            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
            //set the secret key for the tripleDES algorithm
            tdes.Key = keyArray;
            //mode of operation. there are other 4 modes. 
            //We choose ECB(Electronic code Book)

            tdes.Mode = CipherMode.ECB;
            //padding mode(if any extra byte added)
            tdes.Padding = PaddingMode.PKCS7;

            ICryptoTransform cTransform = tdes.CreateDecryptor();
            byte[] resultArray = cTransform.TransformFinalBlock(
                                 toEncryptArray, 0, toEncryptArray.Length);
            //Release resources held by TripleDes Encryptor                
            tdes.Clear();
            //return the Clear decrypted TEXT
            return UTF8Encoding.UTF8.GetString(resultArray);
        }
    }
}
namespace Ov3rExcel
{
    public class GestionExcel
    {
        #region Déclarations
        Excel.Application _ApplicationXL;
        Excel.Workbooks _LesWorkBooks;
        Excel._Workbook _MonClasseur;
        Excel._Worksheet _MaFeuille;
        Excel.Range _MonRange; // le range est une zone selectionnée de façon virutelle ou les calculs s'effectuent. Range est cellule active sont deux chose differentes.
        Excel._Chart _MonGraphique;
        private int processID = 0;
        object _M = System.Reflection.Missing.Value;
        string _NomFichier;

        public enum TypeEncadrement { ToutesCellulesFin, ToutesCellulesMoyen, Aucun, ExtérieurRangeFin, ExtérieurRangeMoyen };
        #endregion

        #region Initialisation et enregistrement
        /// <summary>
        ///  Classe s'interfaçant en une appli et Microsoft.Office.Interop.Excel
        /// </summary>
        public GestionExcel(bool visible = false)
        {
            Init(visible);
        }

        /// <summary>
        ///  Classe s'interfaçant en une appli et Microsoft.Office.Interop.Excel
        /// </summary>
        public GestionExcel(out bool Succes, bool visible = false)
        {
            Succes = string.IsNullOrWhiteSpace(Init(visible));
        }

        private string Init(bool visible)
        {
            foreach (Process item in Process.GetProcessesByName("EXCEL"))
            {
                this.processID -= item.Id;
            }
            try
            {
                //Démarre Excel et récupère l'application
                _ApplicationXL = new Microsoft.Office.Interop.Excel.Application();
                _ApplicationXL.Visible = visible;
                return "";
            }
            catch (Exception e)
            {
                return (e.Message);
            }
            finally
            {
                foreach (Process item in Process.GetProcessesByName("EXCEL"))
                {
                    this.processID += item.Id;
                }
            }
        }

        /// <summary>
        ///  Function créant un classeur vierge
        /// </summary>
        public bool CreerNouveauFichier()
        {
            return string.IsNullOrWhiteSpace(CreerNouveauFichier("Feuil1"));
        }

        /// <summary>
        ///  Function créant un classeur vierge et renommant la feuille active
        /// </summary>
        public string CreerNouveauFichier(string NouveauNomFeuille)
        {
            try
            {
                //Récupère le WorkBook
                _MonClasseur = _ApplicationXL.Workbooks.Add(1);
                //Récupère la feuille Active
                _MaFeuille = (Excel._Worksheet)_MonClasseur.Sheets[1];
                //et la remonne
                _MaFeuille.Name = NouveauNomFeuille;
                //on initialise un Range
                _MonRange = _MaFeuille.get_Range("A1", "A1");
                return "";
            }
            catch (Exception e)
            {
                return e.Message;

            }
        }

        /// <summary>
        /// Teste si le fichier existe, si oui on l'ouvre sinon on le crée et l'enregistre
        /// </summary>
        /// <param name="NomFichier"></param>
        /// <param name="NomFeuille"></param>
        /// <returns></returns>
        public bool CreerOuOuvrirFichier(string NomFichier)
        {
            if (File.Exists(NomFichier))
                return string.IsNullOrWhiteSpace(OuvrirFichierExistant(NomFichier));
            else
            {
                if (!CreerNouveauFichier()) return false;
                if (!string.IsNullOrWhiteSpace(EnregistrerSous(NomFichier))) return false;
                return true;
            }
        }

        /// <summary>
        ///  NomFichier: chemin complet.
        ///  Fonctionne corectement à partir de Office 2000
        /// </summary>
        public string EnregistrerSous(string NomFichier)
        {
            _NomFichier = NomFichier;
            try
            //on enregistre sous
            {
                object FileName = (object)NomFichier;

                _MonClasseur.SaveAs(FileName, _M, _M, _M, _M, _M, Excel.XlSaveAsAccessMode.xlNoChange, _M, _M, _M, _M, _M);
                return "";
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

        /// <summary>
        ///  NomFichier: chemin complet.
        ///  Fonctionne corectement à partir de Office 2000
        /// </summary>
        public string EnregistrerSousUnicode(string NomFichier)
        {
            _NomFichier = NomFichier;
            try
            //on enregistre sous
            {
                object FileName = (object)NomFichier;

                _MonClasseur.SaveAs(FileName, Excel.XlFileFormat.xlUnicodeText, _M, _M, _M, _M, Excel.XlSaveAsAccessMode.xlNoChange, _M, _M, _M, _M, _M);
                return "";
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

        /// <summary>
        ///  NomFichier: chemin complet.
        ///  Feuille de calcul XML 2003.
        /// </summary>
        public string EnregistrerSousFeuilleCalculXML(string NomFichier)
        {
            _NomFichier = NomFichier;
            try
            //on enregistre sous
            {
                object FileName = (object)NomFichier;

                _MonClasseur.SaveAs(FileName, Excel.XlFileFormat.xlXMLSpreadsheet, _M, _M, (object)false, (object)false, Excel.XlSaveAsAccessMode.xlNoChange, _M, _M, _M, _M, _M);
                return "";
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

        /// <summary>
        ///  NomFichier : chemin complet.
        /// </summary>
        public string OuvrirFichierExistant(string NomFichier)
        {
            try
            {
                _LesWorkBooks = _ApplicationXL.Workbooks;
                //ouvrir le fichier Excel désiré 
                _MonClasseur = _LesWorkBooks.Open(NomFichier, _M, _M, _M, _M, _M, _M, _M, _M, _M, _M, _M, _M, _M, _M);

                //Active la feuille 1
                _MaFeuille = (Excel._Worksheet)_MonClasseur.Sheets[1];
                _MaFeuille.Activate();

                //on initialise un Range
                _MonRange = _MaFeuille.get_Range("A1", "A1");
                return "";
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

        /// <summary>
        ///  NomFichier : chemin complet.
        ///  NomFeuille : Active la feuille dans le classeur
        /// </summary>
        public string OuvrirFichierExistant(string NomFichier, string NomFeuille)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(OuvrirFichierExistant(NomFichier))) return "0";

                //Active la page choisie
                _MaFeuille = (Excel._Worksheet)_MonClasseur.Sheets[NomFeuille];
                _MaFeuille.Activate();

                //on initialise un Range
                _MonRange = _MaFeuille.get_Range("A1", "A1");
                return "";
            }
            catch (Exception e)
            {
                return e.Message;
            }

        }

        /// <summary>
        ///  Enregistre le document avec le nom prédéfini.
        /// </summary>
        public string Sauver()
        {
            try
            {
                _MonClasseur.Save();

                return "";
            }
            catch (Exception e)
            {
                return e.Message;
            }

        }

        /// <summary>
        ///  Sauve, ferme et libére les variables
        /// </summary>
        public void FermerWoorkBook()
        {
            this.FermerWoorkBook(true);
        }

        /// <summary>
        /// Ferme et libére les variables.
        /// </summary>
        /// <param name="Sauver">Indique s'il faut sauver avant de fermer.</param>
        public string FermerWoorkBook(bool Sauver)
        {
            try
            {
                //sauve
                if (Sauver) this.Sauver();

                //libére la variable feuille
                if (_MaFeuille != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_MaFeuille);
                    _MaFeuille = null;
                }

                //ferme le WB et libére la variable
                if (_MonClasseur != null)
                {
                    _MonClasseur.Close(true, _NomFichier, _M);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_MonClasseur);
                    _MonClasseur = null;
                }
                return "";
            }
            catch
            {
                return "Impossible de fermer le classeur";
            }
        }

        /// <summary>
        ///  Quitte, sauve, ferme et libére les variables
        /// </summary>
        public void QuitterXl()
        {
            this.QuitterXl(true);
        }

        /// <summary>
        ///  Quitte, ferme et libére les variables
        /// </summary>
        /// <param name="Sauver">Indique s'il faut sauver avant de fermer.</param>
        public void QuitterXl(bool Sauver)
        {
            //ferme le WB
            FermerWoorkBook(Sauver);

            //libére la variable WorkBooks
            if (_LesWorkBooks != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_LesWorkBooks);
                _LesWorkBooks = null;
            }

            //quitte XL et libére la variable
            if (_ApplicationXL != null)
            {
                _ApplicationXL.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_ApplicationXL);
                _ApplicationXL = null;
            }
        }

        #endregion

        #region Ecriture Données

        /// <summary>
        ///  Function transcrivant un tableau de string sur une ligne. Les nombres seront interprétés comme du texte.
        /// </summary>
        public void EcrireLigne(string[] LigneAEcrire, int NumeroLigne)
        {
            int _Col = 1;

            foreach (string Elements in LigneAEcrire)
            {
                //Ecrit cellule par cellule
                _MaFeuille.Cells[NumeroLigne, _Col] = Elements;
                _Col++;
            }
            SelectionRange(NumeroLigne, 1, 1, 1);
        }

        /// <summary>
        ///  Function transcrivant un tableau de objects sur une ligne. Les nombres seront reconnus comme tels.
        /// </summary>
        public void EcrireLigne(object[] LigneAEcrire, int NumeroLigne)
        {
            int _Col = 1;

            foreach (object Elements in LigneAEcrire)
            {
                //Ecrit cellule par cellule
                _MaFeuille.Cells[NumeroLigne, _Col] = Elements;
                _Col++;
            }
            SelectionRange(NumeroLigne, 1, 1, 1);
        }

        /// <summary>
        /// Ecrit la valeur dans la cellule représentée par son adresse de type A1
        /// </summary>
        /// <param name="Valeur"></param>
        /// <param name="A1"></param>
        public void EcrireCellule(object Valeur, string A1)
        {
            int ligne;
            int colonne;
            this.AdresseCellL1C1(A1, out ligne, out colonne);
            this.EcrireCellule(Valeur, ligne, colonne);
        }

        /// <summary>
        /// Ecrit la valeur dans la cellule définie par ses numéros de ligne et de colonne
        /// </summary>
        /// <param name="Valeur">Valeur à écrire</param>
        public void EcrireCellule(object Valeur, int NumeroLigne, int NumeroColonne)
        {
            _MaFeuille.Cells[NumeroLigne, NumeroColonne] = Valeur;
            SelectionRange(NumeroLigne, NumeroColonne, 1, 1);
        }
        //
        //
        //

        public void createHeaders(int row, int col, string htext, string cell1, string cell2, int mergeColumns, Color fondColor, bool font, int size, Color textColor)
        {
            _MaFeuille.Cells[row, col] = htext;
            _MonRange = _MaFeuille.get_Range(cell1, cell2);
            _MonRange.Merge(mergeColumns);


            _MonRange.Interior.Color = fondColor.ToArgb();
            _MonRange.Font.Color = textColor.ToArgb();
            _MonRange.Borders.Color = System.Drawing.Color.Black.ToArgb();
            _MonRange.Font.Bold = font;
            _MonRange.ColumnWidth = size;
        }
        public void createHeaders(int row, int col, string htext, string cell1, string cell2, int mergeColumns, int fondColor, bool font, int size, int textColor)
        {
            _MaFeuille.Cells[row, col] = htext;
            _MonRange = _MaFeuille.get_Range(cell1, cell2);
            _MonRange.Merge(mergeColumns);


            _MonRange.Interior.Color = fondColor;
            _MonRange.Font.Color = textColor;
            _MonRange.Borders.Color = System.Drawing.Color.Black.ToArgb();
            _MonRange.Font.Bold = font;
            _MonRange.ColumnWidth = size;
        }
        #endregion

        #region Sélection

        /// <summary>
        ///  Active et rennome une feuille selon son numéro
        /// </summary>
        public string ChangerDeFeuilleActive(int NumeroFeuille, string NouveauNom)
        {
            try
            {
                _MaFeuille = (Excel._Worksheet)_MonClasseur.Sheets[NumeroFeuille];
                _MaFeuille.Name = NouveauNom;
                _MaFeuille.Activate();

                return "";
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

        /// <summary>
        ///  Active une feuille selon son nom
        /// </summary>
        public string ChangerDeFeuilleActive(string NomFeuille)
        {
            try
            {
                _MaFeuille = (Excel._Worksheet)_MonClasseur.Sheets[NomFeuille];
                _MaFeuille.Activate();

                return "";
            }
            catch (Exception e)
            {
                return e.Message;
            }

        }

        /// <summary>
        ///  Active une feuille selon son numéro
        /// </summary>
        public string ChangerDeFeuilleActive(int NumeroFeuille)
        {
            try
            {
                _MaFeuille = (Excel._Worksheet)_MonClasseur.Sheets[NumeroFeuille];
                _MaFeuille.Activate();

                return "";
            }
            catch (Exception e)
            {
                return e.Message;
            }

        }
        /// <summary>
        /// Ajoute une feuille au début et l'active
        /// </summary>
        public void AjouterFeuille()
        {
            ChangerDeFeuilleActive(1);
            _MonClasseur.Sheets.Add(_M, _M, _M, _M);
            ChangerDeFeuilleActive(1);
        }
        /// <summary>
        /// Ajoute une feuille, l'active et la renomme
        /// </summary>
        /// <param name="NomFeuille"></param>
        public void AjouterFeuille(string NomFeuille)
        {
            AjouterFeuille();
            RenommerFeuille(NomFeuille);
        }
        /// <summary>
        /// Renomme la feuille active
        /// </summary>
        /// <param name="NomFeuille"></param>
        public string RenommerFeuille(string NomFeuille)
        {
            try
            {
                _MaFeuille.Name = NomFeuille;
                return "";
            }
            catch (Exception e)
            {
                if (e.Message == "Impossible de renommer une feuille comme une autre feuille, une bibliothèque d'objets référencée ou un classeur référencé par Visual Basic.")
                    return "Une feuille est déjà nommée ainsi. La feuille ne sera pas renommée.";
                else return e.Message;
            }

        }
        /// <summary>
        /// Supprime la feuille active
        /// </summary>
        public void SupprimerFeuille()
        {
            _MaFeuille.Delete();
        }

        /// <summary>
        /// Deplace la feuille active avant l'index
        /// </summary>
        /// <param name="Index"></param>
        public void DeplacerFeuilleAvant(int Index)
        {
            _MaFeuille.Move(_MonClasseur.Sheets[Index], _M);
        }

        /// <summary>
        ///  Selectionne le range compris entre deux cellules
        /// </summary>
        public void SelectionRange(string CellTypA1, string CellTypB2)
        {
            _MonRange = _MaFeuille.get_Range(CellTypA1, CellTypB2);
        }
        /// <summary>
        /// Sélectionne une cellule
        /// </summary>
        /// <param name="colonne">La lettre de la colonne</param>
        /// <param name="numLigne">Le numéro de ligne</param>
        public void SelectionUneCellule(string colonne, int numLigne)
        {
            string cellule = colonne.ToUpper() + numLigne.ToString();
            _MonRange = _MaFeuille.get_Range(cellule, cellule);
        }
        /// <summary>
        /// Selectionne un Range.
        /// </summary>
        /// <param name="Ligne">Ligne de la cellule Haut Gauche</param>
        /// <param name="Colonne">Colonne de la cellule Haut Gauche</param>
        /// <param name="Hauteur">Nombre de Lignes</param>
        /// <param name="Largeur">Nombres de Colonnes</param>
        public void SelectionRange(int Ligne, int Colonne, int Hauteur, int Largeur)
        {
            string CellTypA1 = this.AdresseCellTypeA1(Ligne, Colonne);
            string CellTypB2 = this.AdresseCellTypeA1(Ligne + Hauteur - 1, Colonne + Largeur - 1);
            _MonRange = _MaFeuille.get_Range(CellTypA1, CellTypB2);
        }
        /// <summary>
        /// Sélectionne la colonne passée en paramètre.
        /// </summary>
        /// <param name="Colonne"></param>
        public void SelectionColonne(string Colonne)
        {
            string MaColonne = string.Format("{0}:{1}", Colonne, Colonne);
            _MonRange = _MaFeuille.get_Range(MaColonne, _M);
        }
        /// <summary>
        /// Sélectionne la ligne passée en paramètre.
        /// </summary>
        /// <param name="Colonne"></param>
        public void SelectionLigne(int Ligne)
        {
            string MaLigne = string.Format("{0}:{1}", Ligne, Ligne);
            _MonRange = _MaFeuille.get_Range(MaLigne, _M);
        }
        /// <summary>
        /// Supprime une sélection et décalle les données suivant le paramètre.
        /// </summary>
        /// <param name="SensDecallageDonnees"></param>
        public void SupprimerRange(Excel.XlDeleteShiftDirection SensDecallageDonnees)
        {
            _MonRange.Delete(SensDecallageDonnees);
        }

        /// <summary>
        /// Selectionne visuellement le Range,
        /// s'il n'est constitué que d'une cellule, celle ci devient la cellule active
        /// (Permet de deplacer la zone visible de la feuille vers la zone de travail).
        /// </summary>
        public void ActiverRange()
        {
            try
            {
                _MonRange.Select();
            }
            catch { }
        }
        /// <summary>
        /// Atteint la derniere cellule utilisée, et retourne son adresse du type A1.
        /// Attention, une cellule vide avec un formattage particulié
        /// ou une cellule contenant une chaine vide est utilisée.
        /// </summary>
        public string AtteindreDerniereCelluleFeuille()
        {
            _MonRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, _M).Select();
            _MonRange = _ApplicationXL.ActiveCell;
            return this.AdresseCellTypeA1(_MonRange.Row, _MonRange.Column);
        }

        /// <summary>
        /// Atteindre la derniére Cellule utilisée d'une colonne.
        /// Incompatible avec un fichiers texte formatté xls.
        /// </summary>
        /// <param name="Colonne">Numéro de la colonne concernée (de 1 à n)</param>
        /// <returns></returns>
        public int AtteindreDerniereCelluleColonne(int Colonne)
        {
            //on part d'en bas et on cherche la premiere cellule utilisée
            string DerniereCelluleFeuille = this.AdresseCellTypeA1(_MaFeuille.Rows.Count, Colonne);
            _MonRange = _MonRange.get_Range(DerniereCelluleFeuille, DerniereCelluleFeuille);

            _MonRange.get_End(Excel.XlDirection.xlUp).Select();
            _MonRange = _ApplicationXL.ActiveCell;
            return _MonRange.Row;
        }

        /// <summary>
        /// Returne l'adresse, du type A1, de la derniere cellule non vide,
        /// en partant d'une cellule dans une direction définie
        /// </summary>
        /// <param name="CelluleDepart"></param>
        /// <param name="Direction"></param>
        /// <returns></returns>
        public string AtteindreDerniereCelluleNonVide(string CelluleDepart, Excel.XlDirection Direction)
        {
            _MonRange = _MonRange.get_Range(CelluleDepart, CelluleDepart);
            _MonRange.get_End(Direction).Select();
            _MonRange = _ApplicationXL.ActiveCell;
            return this.AdresseCellTypeA1(_MonRange.Row, _MonRange.Column);
        }

        public void Imprime()
        {
            object Copies = (object)1;
            object PrinterName = (object)"PDFCreator sur Ne00:";
            object Vrai = (object)true;

            _ApplicationXL.ActivePrinter = PrinterName.ToString();
            _MonClasseur.PrintOut(_M, _M, Copies, _M, PrinterName, _M, Vrai, _M);

        }
        #endregion

        #region Cellule

        /// <summary>
        ///  Convertit les adresses de Cellule type L1C1 en type A1
        /// </summary>
        public string AdresseCellTypeA1(string L1C1)
        {

            int Ligne = 0;
            int Colonne = 0;

            return this.AdresseCellTypeA1(L1C1, out Ligne, out Colonne);
        }

        /// <summary>
        ///  Convertit les adresses de Cellule type L1C1 en type A1
        /// </summary>
        /// <param name="L1C1">Adresse d'entrée</param>
        /// <param name="Ligne"></param>
        /// <param name="Colonne"></param>
        /// <returns></returns>
        public string AdresseCellTypeA1(string L1C1, out int Ligne, out int Colonne)
        {
            int index = L1C1.IndexOf("C");
            Ligne = Convert.ToInt16(L1C1.Substring(1, index - 1));
            Colonne = Convert.ToInt16(L1C1.Substring(index + 1));

            return this.AdresseCellTypeA1(Ligne, Colonne);
        }

        /// <summary>
        ///  Convertit les numéros de ligne et colonne en adresses de Cellule type A1
        /// </summary>
        public string AdresseCellTypeA1(int Ligne, int Colonne)
        {
            if (Ligne < 1 || Colonne < 1) { return null; }
            else
            {
                if (Colonne < 27)
                {
                    char Lettre = new char();
                    //convertit la valeur de colonne en lettre majuscule A = 65
                    Lettre = Convert.ToChar(Colonne + 64);

                    return string.Format("{0}{1}", Lettre.ToString(), Ligne);
                }
                else
                {
                    ////optient la premiere lettre par division la deuxiéme est le reste
                    //int Lettre2;
                    //int Lettre1 = Math.DivRem(Colonne, 26, out Lettre2);

                    ////converti la valeur de colonne et lettre majuscule A = 65 etc
                    //char Lettre11 = new char();
                    //Lettre11 = Convert.ToChar(Lettre1 + 64);
                    //char Lettre22 = new char();
                    //Lettre22 = Convert.ToChar(Lettre2 + 64);

                    //optient la premiere lettre par division la deuxiéme est le reste
                    int Lettre2;
                    int Lettre1 = Math.DivRem((Colonne - 1), 26, out Lettre2);

                    //converti la valeur de colonne et lettre majuscule A = 65 etc
                    //                    char Lettre11 = new char();
                    char Lettre11;
                    Lettre11 = Convert.ToChar(Lettre1 + 64);
                    //                    char Lettre22 = new char();
                    char Lettre22;
                    Lettre22 = Convert.ToChar(Lettre2 + 65);
                    return string.Format("{0}{1}{2}", Lettre11.ToString(), Lettre22.ToString(), Ligne);

                }
            }
        }

        /// <summary>
        /// Retourne l'adresse de type L1C1 depuis une adresse de type A1
        /// </summary>
        /// <param name="A1"></param>
        /// <returns></returns>
        public string AdresseCellL1C1(string A1)
        {
            int toto;
            int tutu;
            return this.AdresseCellL1C1(A1, out toto, out tutu);
        }

        /// <summary>
        /// Retourne l'adresse de type L1C1 depuis une adresse de type A1, les paramétres numéro de ligne et de colonne sont récupérables
        /// </summary>
        /// <param name="A1"></param>
        /// <returns></returns>
        public string AdresseCellL1C1(string A1, out int NumeroLigne, out int NumeroColonne)
        {
            int[] Lettre = { 0, 0 };
            NumeroLigne = 0;
            NumeroColonne = 0;

            for (int i = 0; i < A1.Length; i++)
            {
                char Car = A1[i];

                if ((Car > 64) && (Car < 91))// c'est une lettre
                {
                    Lettre[i] = Convert.ToInt32(Car) - 64;
                }
                else
                {
                    NumeroLigne = Convert.ToInt32(A1.Substring(i));
                    i = A1.Length;
                }
            }

            if (Lettre[1] == 0)
            {
                NumeroColonne = Lettre[0];
            }
            else
            {
                NumeroColonne = 26 * Lettre[0] + Lettre[1];
            }

            return string.Format("L{0}C{1}", NumeroLigne.ToString(), NumeroColonne.ToString());
        }

        #endregion

        #region Formats et mise en Forme

        /// <summary>
        ///  Encadre les cellules d'un range
        /// </summary>
        public void EncradrementRange(TypeEncadrement Encadrement)
        {
            switch (Encadrement)
            {
                case TypeEncadrement.ToutesCellulesFin:
                    _MonRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    //_MonRange.Borders.Weight = XlBorderWeight.xlThin;
                    break;

                case TypeEncadrement.ToutesCellulesMoyen:
                    _MonRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    _MonRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                    break;

                case TypeEncadrement.Aucun:
                    _MonRange.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    break;

                case TypeEncadrement.ExtérieurRangeFin:
                    _MonRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, _M);
                    break;

                case TypeEncadrement.ExtérieurRangeMoyen:
                    _MonRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, _M);
                    break;

            }
        }

        /// <summary>
        /// Fusionne le range sélectionné
        /// </summary>
        public void FusionnerRange()
        {
            _MonRange.Merge(_M);
        }

        /// <summary>
        /// Colore le fond du range sélectnionné
        /// </summary>
        /// <param name="MaCouleur"></param>
        public void CouleurFondRange(Color MaCouleur)
        {
            _MonRange.Interior.Color = MaCouleur.ToArgb();
        }
        /// <summary>
        /// Obtient le code couleur (win32) du fond de la cellule
        /// </summary>
        /// <returns></returns>
        public double ObtenirCouleurFond()
        {
            return this._MonRange.Interior.Color;
        }
        /// <summary>
        /// Obtient le code couleur (Win32) du texte
        /// </summary>
        /// <returns></returns>
        public double ObtenirCouleurText()
        {
            return this._MonRange.Font.Color;
        }
        /// <summary>
        /// Applique une mise en forme conditionnelle au Range
        /// </summary>
        /// <param name="A1"></param>
        /// <param name="Formule"></param>
        public void MiseEnFormeConditionnelle(string Formule, ColorIndex FontColorIndex, ColorIndex BackColorIndex)
        {

            _MonRange.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, _M, Formule, _M);
            int i = _MonRange.FormatConditions.Count;
            _MonRange.FormatConditions[i].Font.ColorIndex = (int)FontColorIndex;
            _MonRange.FormatConditions[i].Interior.ColorIndex = (int)BackColorIndex;
        }

        /// <summary>
        /// Applique une couleur de police et de fond au range
        /// </summary>
        /// <param name="FontColorIndex"></param>
        /// <param name="BackColorIndex"></param>
        public void MiseEnForme(ColorIndex FontColorIndex, ColorIndex BackColorIndex)
        {
            _MonRange.Interior.ColorIndex = BackColorIndex;
            _MonRange.Font.ColorIndex = FontColorIndex;
        }

        /// <summary>
        /// Mets le range en cours au Format %
        /// </summary>
        public void FormatPourcent()
        {
            //object Forma = (object)"0.00%";
            //_MonRange.NumberFormat = Forma;
        }

        #endregion

        #region Graphique

        /// <summary>
        ///  Crée une courbe sur une nouvelle page graphique
        /// </summary>
        public void CreerGraphique(string NomFeuilleDonnees, string CellHautGauche, int NbreLignes, int NbreColonnes, string Titre, string CellSerie, string NomFeuille)
        {
            //Creation d'une nouvelle feuille graphique au début du classeur
            _MonGraphique = (Excel._Chart)_MonClasseur.Sheets.Add(_MonClasseur.Sheets[1], _M, 1, Excel.XlWBATemplate.xlWBATChart);

            //type de graphique: Nuage de points avec ligne droites, sans marqueurs
            _MonGraphique.ChartType = Excel.XlChartType.xlXYScatterLinesNoMarkers;

            //Donne un nom à la feuille
            _MonGraphique.Name = NomFeuille;

            //Selection de la feuille ou sont les données pour ensuit creer le range
            _MaFeuille = (Excel._Worksheet)_MonClasseur.Sheets[NomFeuilleDonnees];

            //le range commence à CelleHautGauche, descend de NbreLignes, est large de NbreColonnes. Si CellSerie n'est pas dans le Range, alors ce sera le nom de la série
            _MonRange = _MaFeuille.get_Range(CellHautGauche, CellSerie).get_Resize(NbreLignes + 1, NbreColonnes);

            //Injecte la premiére série
            _MonGraphique.SetSourceData(_MonRange, Excel.XlRowCol.xlColumns);

            //Si la Série à un nom, c'est automaiquement le titre, on peut alors le modifier.
            if (_MonGraphique.HasTitle) { _MonGraphique.ChartTitle.Text = Titre; }

            //Récupére la série
            //Series _MaSerie = (Series)_MonGraphique.SeriesCollection(0);

            //lui donne un tire
            //_MaSerie.Name = NomSerie;
            //int toto = ((Lines)_MonGraphique.Lines(1)).Count;
            //Line maligne = (Line)_MonGraphique.Lines(0);

            //maligne.Duplicate();

        }

        /// <summary>
        ///  Ajoute une courbe au graphique actif
        /// </summary>
        public void AjouterCourbe(string Nom, string NomFeuilleDonnee, string CelluleHautX, string CelluleBasX, string CelluleHautY, string CelluleBasY)
        {
            //Crée la nouvelle série
            Excel.SeriesCollection _MesSeries = ((Excel.SeriesCollection)_MonGraphique.SeriesCollection(_M));
            Excel.Series _MaSerie = _MesSeries.NewSeries();

            //passe les paramétres
            _MaSerie.Name = Nom;


            //definiyio, des plage de données du type "='Moyennes'!$C$2:$C$5"
            string PlageDonneesX = string.Format("='{0}'!${1}${2}:${3}${4}", NomFeuilleDonnee, CelluleHautX[0], CelluleHautX[1], CelluleBasX[0], CelluleBasX[1]);
            string PlageDonneesY = string.Format("='{0}'!${1}${2}:${3}${4}", NomFeuilleDonnee, CelluleHautY[0], CelluleHautY[1], CelluleBasY[0], CelluleBasY[1]);
            _MaSerie.XValues = PlageDonneesX;
            _MaSerie.Values = PlageDonneesY;
        }

        #endregion

        #region Lecture Données

        /// <summary>
        /// Recherche le texte à l'intérieur des forumules et retourne l'adresse de la cellule du type A1.
        /// La recherche débute par la cellule A1.
        /// Ne trouve pas le resultat d'une concaténation.
        /// </summary>
        /// <param name="TexteATrouver"></param>
        /// <returns></returns>
        public string Chercher(string TexteATrouver)
        {
            //Par déaut la recherche débute à partir de la case A1
            this.SelectionRange("A1", "A1");
            return this.ChercherSuivant(TexteATrouver);
        }

        /// <summary>
        /// Recherche le texte à l'intérieur des forumules et retourne l'adresse de la cellule du type A1.
        /// La recherche débute la cellule A1.
        /// Retourne la valeur de la cellule en paramétre.
        /// </summary>
        /// <param name="TexteATrouver"></param>
        /// <returns></returns>
        public string Chercher(string TexteATrouver, out string Contenu)
        {
            int Ligne = 0;
            int Colonne = 0;

            string Cellule = this.Chercher(TexteATrouver);
            Contenu = this.LireCellule(Cellule, out  Ligne, out Colonne);
            return Cellule;
        }

        /// <summary>
        /// Recherche le texte à l'intérieur des forumules et retourne l'adresse de la cellule du type A1.
        /// La recherche débute aprés le Range Actif, ce qui est parfois générateur de bug.
        /// </summary>
        /// <param name="TexteATrouver"></param>
        /// <returns></returns>
        public string ChercherSuivant(string TexteATrouver)
        {
            object What = (object)TexteATrouver;

            if (_MonRange.Cells.Count != 1) this.SelectionRange("A1", "A1");
            //fonction find
            _MonRange = _MaFeuille.Cells.Find(What, _MonRange, _M, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, (object)false, _M, (object)false);

            if (_MonRange == null)
            {
                return "Recherche infructeuse";
            }
            else
            {
                //Active la céllule trouvée
                _MonRange.Activate();
                return string.Format("L{0}C{1}", _MonRange.Row.ToString(), _MonRange.Column.ToString());
            }
        }

        /// <summary>
        /// Recherche le texte à l'intérieur des forumules et retourne l'adresse de la cellule du type A1.
        /// La recherche débute aprés le Range.
        /// Retourne la valeur de la cellule en paramétre.
        /// </summary>
        /// <param name="TexteATrouver"></param>
        /// <returns></returns>
        public string ChercherSuivant(string TexteATrouver, out string Contenu)
        {
            int Ligne = 0;
            int Colonne = 0;

            string Cellule = this.ChercherSuivant(TexteATrouver);
            Contenu = this.LireCellule(Cellule, out  Ligne, out Colonne);
            return Cellule;
        }

        /// <summary>
        /// Remplace une chaine dans le range en cours, sans condition de casse, de texte entier ni de format
        /// </summary>
        /// <param name="TexteAChercher">Texte à trouver</param>
        /// <param name="TexteAEcrire">Texte de remplacement</param>
        public void Remplacer(string TexteAChercher, string TexteAEcrire)
        {
            _MonRange.Replace(TexteAChercher, TexteAEcrire, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, _M, false, false);
        }

        /// <summary>
        /// Remplace plusieurs chaines dans le range en cours, sans condition de casse, de texte entier ni de format
        /// </summary>
        /// <param name="TexteAChercher">Tableau de string comportant toutes les occurences à remplacer</param>
        /// <param name="TexteAEcrire">Texte de remplacement</param>
        public void Remplacer(string[] TexteAChercher, string TexteAEcrire)
        {
            foreach (string Texte in TexteAChercher)
                this.Remplacer(Texte, TexteAEcrire);
        }

        /// <summary>
        /// Retourne le contenu de la cellule (Formule, long pour les dates etc...)
        /// </summary>
        /// <param name="AdresseA1"></param>
        /// <returns></returns>
        public string LireCellule(string AdresseA1)
        {
            this.SelectionRange(AdresseA1, AdresseA1);
            return (string)_MonRange.Value2;
        }

        /// <summary>
        /// Retourne le contenu de la cellule (Formule, long pour les dates etc...)
        /// </summary>
        /// <param name="AdresseA1"></param>
        /// <param name="Texte">Retourne le Texte de la Cellule</param>
        /// <returns></returns>
        public string LireCellule(string AdresseA1, out string Texte)
        {
            this.SelectionRange(AdresseA1, AdresseA1);
            Texte = (string)_MonRange.Text;
            return (string)_MonRange.Value2;
        }

        /// <summary>
        /// Retourne le contenu de la cellule.
        /// </summary>
        /// <param name="AdresseA1"></param>
        /// <returns></returns>
        public string LireCellule(string AdresseL1C1, out int NumeroLigne, out int NumeroColonne)
        {
            string AdresseA1 = this.AdresseCellTypeA1(AdresseL1C1, out NumeroLigne, out NumeroColonne);
            return this.LireCellule(AdresseA1);

        }

        /// <summary>
        /// Retourne le contenu de la cellule.
        /// </summary>
        /// <param name="AdresseA1"></param>
        /// <returns></returns>
        public string LireCellule(int NumeroLigne, int NumeroColonne, out string Texte)
        {
            string AdresseA1 = this.AdresseCellTypeA1(NumeroLigne, NumeroColonne);
            return this.LireCellule(AdresseA1, out Texte);
        }

        /// <summary>
        /// Retourne le contenu de la cellule.
        /// </summary>
        /// <param name="AdresseA1"></param>
        /// <param name="Texte">Retourne le Texte de la Cellule</param>
        /// <returns></returns>
        public string LireCellule(int NumeroLigne, int NumeroColonne)
        {
            string AdresseA1 = this.AdresseCellTypeA1(NumeroLigne, NumeroColonne);
            return this.LireCellule(AdresseA1);
        }

        /// <summary>
        /// Protège la feuille, sans mot de passe
        /// </summary>
        public string ProtegerFeuille()
        {
            object Vrai = (object)true;
            object Faux = (object)false;

            try
            {
                _MaFeuille.Protect(_M, Vrai, Vrai, Vrai, _M, _M, _M, _M, _M, _M, _M, _M, _M, _M, _M, _M);
                return "";
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

        /// <summary>
        /// Protège la feuille avec mot de passe
        /// </summary>
        /// <param name="MotDePasse">si chaine vide est saisi, il n'y aura pas de mot de passe</param>
        public string ProtegerFeuille(string MotDePasse)
        {
            object Vrai = (object)true;
            object Faux = (object)false;
            object Password = (object)MotDePasse;

            try
            {
                _MaFeuille.Protect(Password, Vrai, Vrai, Vrai, _M, _M, _M, _M, _M, _M, _M, _M, _M, _M, _M, _M);
                return "";
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

        /// <summary>
        /// Ote la protection d'une feuille, sans mot de passe
        /// </summary>
        public string OterProtectionFeuille()
        {

            try
            {
                _MaFeuille.Unprotect(_M);
                return "";
            }
            catch (Exception e)
            {
                return e.Message;
            }

        }

        /// <summary>
        /// Ote la protection d'une feuille avec mot de passe
        /// </summary>
        /// <param name="MotDePasse"></param>
        public string OterProtectionFeuille(string MotDePasse)
        {
            object Password = (object)MotDePasse;

            try
            {
                _MaFeuille.Unprotect(MotDePasse);
                return "";
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

        /// <summary>
        /// Protège le classeur, sans mot de passe
        /// </summary>
        /// <param name="Structure">La structure est protégée</param>
        /// <param name="Fenetres">Les fenêtre sont protégées</param>
        public string ProtegerClasseur(bool Structure, bool Fenetres)
        {
            object _Structure = (object)Structure;
            object _Fenetres = (object)Fenetres;

            try
            {
                _MonClasseur.Protect(_M, _Structure, _Fenetres);
                return "";
            }
            catch (Exception e)
            {
                return e.Message;
            }

        }

        /// <summary>
        /// Protège le classeur, avec mot de passe
        /// </summary>
        /// <param name="Structure">La structure est protégée</param>
        /// <param name="Fenetres">Les fenêtre sont protégées</param>
        public string ProtegerClasseur(string MotDePasse, bool Structure, bool Fenetres)
        {
            object Password = (object)MotDePasse;
            object _Structure = (object)Structure;
            object _Fenetres = (object)Fenetres;

            try
            {
                _MonClasseur.Protect(_M, _Structure, _Fenetres);
                return "";
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

        /// <summary>
        /// Ote la protection d'un classeur
        /// </summary>
        public string OterProtectionClasseur()
        {
            try
            {
                _MonClasseur.Unprotect(_M);
                return "";
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

        /// <summary>
        /// Ote la protection d'un classeur avec mot de passe
        /// </summary>
        public string OterProtectionClasseur(string MotDePasse)
        {
            object Password = (object)MotDePasse;

            try
            {
                _MonClasseur.Unprotect(MotDePasse);
                return "";
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

        #endregion

        public void quitExcel()
        {
            Process.GetProcessById(processID).Kill();
        }
    }
    public enum ColorIndex
    {
        Jaune = 6,
        Mauve = 7,
        Orange = 44,
        Vert = 4,
        Azur = 8,
        Rouge = 3,
        Noir = 1,
        Automatique = 0
    }
}
namespace Ov3rStatic
{
    static class MyClass
	{
        /// <summary>
        /// Effectue une sauvegarde de la base de données
        /// </summary>
        /// <param name="cn">Instance de la connexion</param>
        /// <param name="database">Nom de la base à sauvegarder</param>
        /// <param name="pathDatabase">Emplacement de la sauvegarde</param>
        /// <param name="timeOut">Temps limite de la requête</param>
        /// <returns>Return une chaine vide si réquête réussite sinon retourne le message erreur</returns>
        public static string SaveDatabase(SqlConnection cn, string database, string pathDatabase, int timeOut = 1000000)
        {
            try
            {
                cn.Open();
                string req = "USE " + database + ";";
                req += "BACKUP DATABASE " + database;
                req += @" TO DISK = '" + pathDatabase + @"\" + database + ".bak'";
                req += " WITH FORMAT,";
                req += "MEDIANAME = 'Z_SQLServerBackups',";
                req += "NAME = 'Full Backup of " + database + "';";
                cn.Open();

                SqlCommand requete = new SqlCommand(req, cn);
                requete.CommandTimeout = timeOut;
                requete.ExecuteNonQuery();

                return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                cn.Close();
            }

        }
        /// <summary>
        /// Effectue une sauvegarde de la base de données
        /// </summary>
        /// <param name="cn">Chaine de la connexion</param>
        /// <param name="database">Nom de la base à sauvegarder</param>
        /// <param name="pathDatabase">Emplacement de la sauvegarde</param>
        /// <param name="timeOut">Temps limite de la requête</param>
        /// <returns>Return une chaine vide si réquête réussite sinon retourne le message erreur</returns>
        public static string SaveDatabase(string chainCn, string database, string pathDatabase, int timeOut = 1000000)
        {
            SqlConnection cn = new SqlConnection(chainCn);            
            try
            {
                cn.Open();
                string req = "USE " + database + ";";
                req += "BACKUP DATABASE " + database;
                req += @" TO DISK = '" + pathDatabase + @"\" + database + ".bak'";
                req += " WITH FORMAT,";
                req += "MEDIANAME = 'Z_SQLServerBackups',";
                req += "NAME = 'Full Backup of " + database + "';";
                cn.Open();

                SqlCommand requete = new SqlCommand(req, cn);
                requete.CommandTimeout = timeOut;
                requete.ExecuteNonQuery();

                return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                cn.Close();
            }
        }
	}
    
}
