using BlazorInputFile;
using System;
using System.IO;


namespace A2B_App.Shared.Sox
{
    public class FileLog
    {
        public static void Write(string msg, string filename)
        {
            string dtToday = DateTime.Now.ToString("yyyyMMdd");
            string dateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            string startupPath = Environment.CurrentDirectory;
            string path = Path.Combine(startupPath, "include", filename + "_" + dtToday + "_log.txt");
            if (!File.Exists(path))
            {
                File.Create(path).Dispose();
            }
            using (StreamWriter sw = File.AppendText(path))
            {
                sw.Write($"[{dateTime}] ==> {msg}" + Environment.NewLine);
            }
        }

        
    }

    public class WriteLog 
    {
        public void Display(object someObject)
        {
            var jsonData = Newtonsoft.Json.JsonConvert.SerializeObject(someObject);
            System.Diagnostics.Debug.WriteLine(jsonData);
        }
    }

    public class FileUpload
    {
        public IFileListEntry IFileEntry { get; set; }
        public int Position { get; set; }
    }


}
