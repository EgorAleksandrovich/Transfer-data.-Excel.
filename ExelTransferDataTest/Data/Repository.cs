using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExelTransferDataTest.Data
{
    public class Repository
    {
        private DataSet _ds;

        private static string GetConnectionString(string file)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            string extension = file.Split('.').Last();

            if (extension == "xls")
            {
                //Excel 2003 and Older
                props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
                props["Extended Properties"] = "Excel 8.0";
            }
            else if (extension == "xlsx")
            {
                //Excel 2007, 2010, 2012, 2013
                props["Provider"] = "Microsoft.ACE.OLEDB.12.0";
                props["Extended Properties"] = "Excel 12.0 XML";
            }
            else
                throw new Exception(string.Format("error file: {0}", file));

            props["Data Source"] = file;

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }

        private static DataSet GetDataSetFromExcelFile(string file)
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
            }
            return ds;
        }

        public void CreateExcelFile(string sourcePath, string targetPath, string newFileName)
        {
            foreach (string existingFileName in Directory.GetFiles(sourcePath, "*.xlsx").Select(Path.GetFileName))
            {
                if (newFileName == existingFileName)
                {
                    throw new System.ArgumentException(string.Format("Файл в папке" + targetPath + " с именем " + newFileName + " уже существует!"));
                }
                else
                {
                    if (newFileName == null)
                    {
                        throw new ArgumentNullException("Не указано имя файла в которое будет скопированы данные!");
                    }
                    else
                    {
                        string tempFileName = "empty.xlsx";
                        string sourceFile = System.IO.Path.Combine(sourcePath, tempFileName);
                        string destFile = System.IO.Path.Combine(targetPath, newFileName);
                        if (!System.IO.Directory.Exists(targetPath))
                        {
                            System.IO.Directory.CreateDirectory(targetPath);
                        }
                        System.IO.File.Copy(sourceFile, destFile, true);
                    }
                }
            }
        }
    }
}
