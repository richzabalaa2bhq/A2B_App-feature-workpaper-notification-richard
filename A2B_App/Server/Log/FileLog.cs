using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace A2B_App.Server.Log
{
    public class FileLog
    {
        //public static void Write(string msg, string filename)
        //{
        //    string dtToday = DateTime.Now.ToString("yyyyMMdd");
        //    string dateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
        //    string startupPath = Environment.CurrentDirectory;
        //    string path = Path.Combine(startupPath, "include", filename + "_" + dtToday + "_log.txt");
        //    if (!File.Exists(path))
        //    {
        //        File.Create(path).Dispose();
        //    }
        //    using (StreamWriter sw = File.AppendText(path))
        //    {
        //        sw.Write($"[{dateTime}] ==> {msg}" + Environment.NewLine);
        //    }
        //}
    }
}
