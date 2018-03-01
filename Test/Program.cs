using ExelTransferDataTest.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Test
{
    class Program
    {
        static private Repository _repository;
        static private DataSet _ds;
        static private DataTable _dt;
        static private Helpers _helpers;
        static private string _oldFile;
        static private string _newFile;
        static List<string> _thicknessOfMaterials;
        static private int _count = 4;
        static string _sql = null;


        static void Main(string[] args)
        {
            _dt = new DataTable();
            _ds = new DataSet();
            _repository = new Repository();
            _helpers = new Helpers();
            _oldFile = @"E:\Development\Projects\Transfer data exel\1.xlsx";
            _newFile = @"E:\Development\Projects\Transfer data exel\empty.xlsx";
            _ds = _repository.GetDataSetFromExcelFile(_oldFile);
            _helpers.Ds = _ds;
            _thicknessOfMaterials = _helpers.GetThicknessOfMaterials();
            System.Data.OleDb.OleDbCommand myCommand = new System.Data.OleDb.OleDbCommand();
            OleDbConnection MyConnection = new OleDbConnection(_repository.GetConnectionString(_newFile));
            MyConnection.Open();
            myCommand.Connection = MyConnection;

            foreach (string thkness in _thicknessOfMaterials)
            {
                _dt = _helpers.GetAllPartWithThikness(thkness);

                myCommand.Parameters.Add(new OleDbParameter("@thikness", string.Format("Материал с толщиной {0}мм", thkness.ToString())));
                _sql = "insert into [Распил$C:C" + _count + "]  values (@thikness)";
                myCommand.CommandText = _sql;
                myCommand.ExecuteNonQuery();
                myCommand.Parameters.Clear();
                _count++;

                foreach (DataRow dr in _dt.Rows)
                {
                    _sql = "insert into [Распил$A:J" + _count + "]  values (@sectionNumber, @numberOfDetail, @nameOfDitails, '', '', '', '', @x, @y, @number)";
                    myCommand.CommandText = _sql;
                    myCommand.Parameters.Add(new OleDbParameter("@sectionNumber", dr["SectionNumber"].ToString()));
                    myCommand.Parameters.Add(new OleDbParameter("@numberOfDetail", _count.ToString()));
                    myCommand.Parameters.Add(new OleDbParameter("@nameOfDitails", dr["NameOfDitails"].ToString()));
                    myCommand.Parameters.Add(new OleDbParameter("@x", dr["X"].ToString()));
                    myCommand.Parameters.Add(new OleDbParameter("@y", dr["Y"].ToString()));
                    myCommand.Parameters.Add(new OleDbParameter("@number", dr["Number"].ToString()));
                    myCommand.ExecuteNonQuery();
                    myCommand.Parameters.Clear();
                    _count++;
                }
                _count += 2;
            }
            MyConnection.Close();
        }
    }
}
