using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExelTransferDataTest.Data
{
    public class Helpers
    {
        private DataSet _ds;
        private int _detailCount;

        private List<object> GetThicknessOfMaterials()
        {
            List<object> thicknessOfMaterials = new List<object>();
            thicknessOfMaterials = GetValuesFromRows(4);
            thicknessOfMaterials = RemoveNullOrEmptyRow(thicknessOfMaterials);
            return thicknessOfMaterials;
        }

        private List<object> GetNumberSections()
        {
            List<object> numberSections = new List<object>();
            numberSections = GetValuesFromRows(0);
            numberSections = RemoveNullOrEmptyRow(numberSections);
            return numberSections;
        }

        public List<object> GetValuesFromRows(int rowNumber)
        {
            return _ds.Tables[0].AsEnumerable().Select(r => r[rowNumber]).Distinct().ToList();
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
            newTable.Columns.Add("Y", typeof(int));
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
                    newNumberSection = Convert.ToInt32(_ds.Tables[0].Rows[i - 1][0]);
                    if (newNumberSection != curentNumberSection)
                    {
                        curentNumberSection = newNumberSection;
                        sectionName = "C" + newNumberSection;
                    }
                }
                if (thikness == _ds.Tables[0].Rows[i][4].ToString())
                {
                    _detailCount += 1;
                    newTable.Rows.Add(sectionName, _detailCount, GetNameOfDetail(i, sectionName, numberOfRowWithNumberSection), GetX(i), GetY(i), GetNumberOfDetails(i), thikness);
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
            string pattern = @"[\.\,]+[\d]{0,}";
            Regex rgx = new Regex(pattern);
            string x = rgx.Replace(_ds.Tables[0].Rows[rowNumber][2].ToString(), "");
            return Convert.ToInt32(x);
        }

        public int GetY(int rowNumber)
        {
            string pattern = @"[\.\,]+[\d]{0,}";
            Regex rgx = new Regex(pattern);
            string y = rgx.Replace(_ds.Tables[0].Rows[rowNumber][2].ToString(), "");
            return Convert.ToInt32(y);
        }

        public int GetNumberOfDetails(int rowNumber)
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
