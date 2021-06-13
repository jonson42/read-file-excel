using ExcelDataReader;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace WindowsFormsControlLibrary1
{
    public class Excell
    {
        private string _path { get; set; }
        private List<object> _list { get; set; }
        public Excell(string path)
        {
            _path = path;
            _list = new List<object>();
        }
        public List<object> ReadFile(string sheetName)
        {
            using (var stream = System.IO.File.Open(_path,FileMode.Open, FileAccess.Read))
            {

                IExcelDataReader excelDataReader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

                var conf = new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = a => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };

                DataSet dataSet = excelDataReader.AsDataSet(conf);
                DataRowCollection row = dataSet.Tables[sheetName].Rows;
                List<object> rowDataList = null;
                List<object> allRowsList = new List<object>();
                foreach (DataRow item in row)
                {
                    rowDataList = item.ItemArray.ToList();
                    allRowsList.Add(rowDataList);
                }
                _list = allRowsList;
                return allRowsList;
            }
        }
    }
}
