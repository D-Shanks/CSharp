using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace test4
{
    class writeMsg
    {
        public static string Path;
        public static int count;
        public string[] sid(string fpath)
        {
            fileRead fr = new fileRead();
            DirectoryInfo dir = new DirectoryInfo(fpath);
            FileSystemInfo[] fsi = dir.GetFileSystemInfos();
            int i = 0;
            string[] sid = new string[fsi.Length];
            foreach(FileSystemInfo file in fsi)
            {
                //fname = fr.fileName(file.ToString());
                sid[i] = fr.fileName(file.ToString())[0];
                i++;  //i+1=考生数量
            }
            count = i;
            return sid;
        }

        public string[] sname(string fpath)
        {
            fileRead fr = new fileRead();
            DirectoryInfo dir = new DirectoryInfo(fpath);
            FileSystemInfo[] fsi = dir.GetFileSystemInfos();
            int i = 0;
            string[] sname = new string[fsi.Length];
            foreach (FileSystemInfo file in fsi)
            {
                sname[i] = fr.fileName(file.ToString())[1];
                i++;  //i+1=考生数量
            }
            return sname;
        }
        //获取路径
        public void road(string path)
        {
            string spath = path;
            Path = spath;
        }
    }
}
