using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;

namespace ExelTransferDataTest.Data
{
    public class Repository
    {
        private string _targetFile;

        public string GetConnectionString(string file)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            string extension = file.Split('.').Last();

            if (extension == "xls")
            {
                //Excel 2003 and Older
                props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
                props["Extended Properties"] = "Excel 8.0";
                //Excel 8.0;HDR=NO;IMEX=3;READONLY=FALSE\
            }
            else if (extension == "xlsx")
            {
                //Excel 2007, 2010, 2012, 2013
                props["Provider"] = "Microsoft.ACE.OLEDB.12.0";
                props["Extended Properties"] = "Excel 12.0 Xml;";
                //;Extended Properties=\"Excel 12.0 Xml; HDR=YES\";"
            }
            else
            {
                MessageBox.Show(string.Format("Файл не найден: {0}", file));
                return null;
            }

            props["Data Source"] = file;

            StringBuilder connectionString = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                connectionString.Append(prop.Key);
                connectionString.Append('=');
                connectionString.Append(prop.Value);
                connectionString.Append(';');
            }
            return connectionString.ToString();
        }

        public DataSet GetDataSetFromExcelFile(string file)
        {
            DataSet ds = new DataSet();
            string connectionString = GetConnectionString(file);

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data
                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();

                    if (!sheetName.EndsWith("$"))
                        continue;

                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                    DataTable dt = new DataTable();
                    dt.TableName = sheetName;

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);

                    ds.Tables.Add(dt);
                }
                cmd = null;
                conn.Close();
                return ds;
            }
        }

        public string CreateExcelFile(string sourceFile, string targetFilePath, string newFileName)
        {
            if (newFileName == null)
            {
                MessageBox.Show("Не указано имя файла в которое будет скопированы данные!");
                return null;
            }
            else
            {
                newFileName += ".xlsx";

                if (MatchFailNamesInFolder(targetFilePath, newFileName))
                {
                    MessageBox.Show(string.Format("Файл в папке \"" + targetFilePath + "\" с именем \"" + newFileName + "\" уже существует! Введине другое имя!"));
                    return null;
                }
                else
                {
                    _targetFile = System.IO.Path.Combine(targetFilePath, newFileName);
                    if (!System.IO.Directory.Exists(targetFilePath))
                    {
                        System.IO.Directory.CreateDirectory(targetFilePath);
                    }
                    System.IO.File.Copy(sourceFile, _targetFile, true);
                }
            }
            return _targetFile;
        }

        public bool MatchFailNamesInFolder(string targetPathFolder, string targetFileName)
        {
            foreach (string existingFileName in Directory.GetFiles(targetPathFolder, "*.xlsx").Select(Path.GetFileName))
            {
                if (existingFileName == targetFileName)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
