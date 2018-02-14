using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExelTransferDataTest.Data
{
    public class Repository
    {
        private DataSet _ds;
        private int _detailCount;

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

        private List<object> GetThicknessOfMaterials()
        {
            List<object> thicknessOfMaterials = new List<object>();
            thicknessOfMaterials = GetValuesFromRows(_ds, 4);
            thicknessOfMaterials = RemoveNullOrEmptyRow(thicknessOfMaterials);
            return thicknessOfMaterials;
        }

        private List<object> GetNumberSections()
        {
            List<object> numberSections = new List<object>();
            numberSections = GetValuesFromRows(_ds, 0);
            numberSections = RemoveNullOrEmptyRow(numberSections);
            return numberSections;
        }

        public static List<object> GetValuesFromRows(DataSet ds, int rowNumber)
        {
            return ds.Tables[0].AsEnumerable().Select(r => r[rowNumber]).Distinct().ToList();
        }

        public static List<object> RemoveNullOrEmptyRow(List<object> list)
        {
            foreach (Object o in list)
            {
                if (string.IsNullOrWhiteSpace(o.ToString()))
                {
                    list.Remove(o);
                    break;
                }
            }
            return list;
        }

        private string GetProjectName()
        {
            string projectName;
            projectName = _ds.Tables[0].Rows[0][1].ToString().Split('.').First();
            return projectName;
        }

        public DataTable GetAllPartWithThikness(string thikness)
        {
            DataTable newTable = new DataTable();
            newTable.Columns.Add("SectionNumber", typeof(string));
            newTable.Columns.Add("Position", typeof(string));
            newTable.Columns.Add("NameOfDitails", typeof(string));
            newTable.Columns.Add("X", typeof(int));
            newTable.Columns.Add("X", typeof(int));
            newTable.Columns.Add("Number", typeof(int));
            newTable.Columns.Add("Thikness", typeof(int));
            string sectionName = "C1";
            int curentNumberSection = 1;
            int newNumberSection = 0;
            int numberOfRowWithNumberSection = 1; 

            int rowNumbers = _ds.Tables[0].Rows.Count;

            for (int i = 1; i <= rowNumbers; i++)
            {
                if (!string.IsNullOrWhiteSpace(_ds.Tables[0].Rows[i - 1][0].ToString()))
                {
                    numberOfRowWithNumberSection = i - 1;
                    newNumberSection = Convert.ToInt32(_ds.Tables[0].Rows[i-1][0]);
                    if (newNumberSection != curentNumberSection)
                    {
                        curentNumberSection = newNumberSection;
                        sectionName = "C" + newNumberSection;
                    }
                }
                if(thikness == _ds.Tables[0].Rows[i][4].ToString())
                {
                    _detailCount += 1;
                    newTable.Rows.Add(sectionName, _detailCount, GetNameOfDetail(i, sectionName, numberOfRowWithNumberSection), GetX(i), GetY(i), GetNumber(i), thikness);
                }
            }
            return null;
        }

        public string GetNameOfDetail(int rowNumber, string sectionName, int numberOfRowWithNumberSection)
        {
            string pattern = "\\s{0,}" + GetRemoveString(numberOfRowWithNumberSection) + ".";
            Regex rgx = new Regex(pattern);
            string nameDetail = rgx.Replace(_ds.Tables[0].Rows[rowNumber][1].ToString(), " ");
            return nameDetail;
        }

        public int GetX(int rowNumber)
        {
            return Convert.ToInt32(_ds.Tables[0].Rows[rowNumber][2].ToString());
        }

        public int GetY(int rowNumber)
        {
            return Convert.ToInt32(_ds.Tables[0].Rows[rowNumber][3].ToString());
        }

        public int GetNumber(int rowNumber)
        {
            return Convert.ToInt32(_ds.Tables[0].Rows[rowNumber][5].ToString());
        }

        private string GetRemoveString(int rowNumberWithNumberSection)
        {
            string removeString;
            removeString = _ds.Tables[0].Rows[rowNumberWithNumberSection][1].ToString();
            return removeString;
        }
    }
}
