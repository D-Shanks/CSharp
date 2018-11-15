using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace test4
{
    class fileRead
    {
        public static List<string> id;
        public static List<string> name;
        public static List<int> score;
        public void readfile()
        {
            FileStream F = new FileStream("d:\\result.dat",
                FileMode.Open, FileAccess.Read);

            F.Position = 0;
            //读版本号
            byte[] istrlength = new byte[4];
            F.Read(istrlength, 0, 4);
            int strlength = BitConverter.ToInt32(istrlength, 0);
            byte[] strArrayread = new byte[strlength];
            F.Read(strArrayread, 0, strlength);
            string strRead = System.Text.Encoding.Default.GetString(strArrayread);
            //读考试要求
            byte[] irelength = new byte[4];
            F.Read(irelength, 0, 4);
            int relength = BitConverter.ToInt32(irelength, 0);
            byte[] reArrayread = new byte[relength];
            F.Read(reArrayread, 0, relength);
            string reRead = System.Text.Encoding.Default.GetString(reArrayread);
            //读考生数量
            byte[] cbuffer = new byte[4];
            F.Read(cbuffer, 0, cbuffer.Length);
            int cread = BitConverter.ToInt32(cbuffer, 0);
            //读考生信息
            writeMsg wm = new writeMsg();
            int n = writeMsg.count;
            int[] len = new int[n];
            int[] num = new int[n];
            string[] namestr = new string[n];
            string[] idstr = new string[n];
            int[] scoreint = new int[n];
            for (int a = 0; a < n; a++)
            {
                //总字节数
                byte[] numbuffer = new byte[4];
                F.Read(numbuffer, 0, numbuffer.Length);
                num[a] = BitConverter.ToInt32(numbuffer, 0);
                //ID
                byte[] iidstr = new byte[4];
                F.Read(iidstr, 0, 4);
                int lenid = BitConverter.ToInt32(iidstr, 0);
                byte[] idArray = new byte[lenid];
                F.Read(idArray, 0, lenid);
                idstr[a] = System.Text.Encoding.Default.GetString(idArray);
                //name
                byte[] inamestr = new byte[4];
                F.Read(inamestr, 0, 4);
                int lenna = BitConverter.ToInt32(inamestr, 0);
                byte[] nameArray = new byte[lenna];
                F.Read(nameArray, 0, lenna);
                namestr[a] = System.Text.Encoding.Default.GetString(nameArray);
                //score
                byte[] iscore = new byte[4];
                F.Read(iscore, 0, iscore.Length);
                scoreint[a] = BitConverter.ToInt32(iscore, 0);
            }
            
            id = new List<string>(idstr);
            name = new List<string>(namestr);
            score = new List<int>(scoreint);
            F.Close();
        }
        //字符串分割
        public string[] fileName(string filename)
        {
            string[] str = filename.Split('_','.');
            List<string> s = new List<string>(str);
            return str;
        }
    }
}
