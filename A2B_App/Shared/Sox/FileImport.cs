using Microsoft.AspNetCore.Http;
using System.Collections.Generic;

namespace A2B_App.Shared.Sox
{
    public class FileImport
    {
        public IFormFile File {get; set; }
        //0 - Rcm
        //1 - SoxTracker
        //2 - SampleSelection
        //3 - KeyReport
        //4 - Sod
        public int Process { get; set; }
        public string ClientName { get; set; }
    }

    public class ImportFields
    {
        public List<ColumnVal> ListExcelColumns { get; set; }
        public List<DBColumnVal> ListDatabaseColumns { get; set; }
        public string Filename { get; set; }
    }

   

    public class ColumnVal
    {
        public int Index { get; set; }
        public string ExcelColumnName { get; set; }
    }

    public class DBColumnVal
    {
        public int Position { get; set; }
        public ColumnVal ExcelColumn { get; set; }
        public string DbColumnName { get; set; }
    }

    public class RowVal 
    { 
        public int Row { get; set; }
        public List<ColumnVal> Column { get; set; }
    } 


    public enum Process
    {
        Rcm = 0,
        SoxTracker = 1,
        SampleSelection = 2,
        KeyReport = 3,
        Sod = 4
    }
}
