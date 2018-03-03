using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using ExelTransferDataTest.Data;
using System.Data;
using System.IO;

namespace ExelTransferDataTest.ViewModel
{
    class ViewModel : ViewModelBase
    {
        static private Repository _repository;
        static private Helpers _helpers;
        private string _selectedFile;
        private string _newFile;
        private DataSet _ds;
        private DataTable _dt;
        List<string> _thicknessOfMaterials;
        private int _count = 4;
        private IOService _ioService;
        string _sql = null;

        public ViewModel()
        {
            _dt = new DataTable();
            _ds = new DataSet();
            _repository = new Repository();
            _helpers = new Helpers();
            TransferCommand = new RelayCommand(TransferData);

            OpenCommand = new RelayCommand(OpenFile);
        }

        public ViewModel(IOService ioService)
        {
            _ioService = ioService;
        }

        public void TransferData()
        {
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\empty.xlsx");
            _ds = _repository.GetDataSetFromExcelFile(_selectedFile);
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

        public RelayCommand TransferCommand { get; set; }
        public RelayCommand OpenCommand { get; set; }

        public string SelectedFile
        {
            get
            {
                return _selectedFile;
            }
            set
            {
                _selectedFile = value;
                OnPropertyChanged("SelectedFile");
            }
        }

        public string NewFileName
        {
            get
            {
                return _newFile;
            }
            set
            {
                _newFile = value;
                OnPropertyChanged("NewFileName");
            }
        }
        private void OpenFile()
        {
            SelectedFile = _ioService.OpenFileDialog(@"c:\Where\My\File\Usually\Is.txt");
            if (SelectedFile == null)
            {
                SelectedFile = string.Empty;
            }
        }
    }
}
