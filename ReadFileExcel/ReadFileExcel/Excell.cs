using ExcelDataReader;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

    public class Excell
    {
        private string _path { get; set; }
        private List<object> _list { get; set; }
        public Excell(string path)
        {
            _path = path;
            _list = new List<object>();
        }
        public List<List<string>> ReadFile(string sheetName)
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
            var temp = JsonConvert.SerializeObject(allRowsList);
                return JsonConvert.DeserializeObject<List<List<string>>>(temp); 
            }
        }
    }
