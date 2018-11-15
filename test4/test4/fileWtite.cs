using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace test4
{
    class fileWtite
    {
        public void writefile()
        {

            FileStream F = new FileStream("d:\\result.dat",
                FileMode.Truncate, FileAccess.ReadWrite);

            //int i = 2018;
            //string str = "成都大学CDDX";
            //初始化Readword 
            ReadWord rw = new ReadWord();
            //初始化writeMsg
            writeMsg wm = new writeMsg();
            LCS<string> lcs = new LCS<string>(rw.readList(), rw.readList1());
            LCS<string> lcs1 = new LCS<string>(rw.readList(), rw.readList2());
            LCS<string> lcs2 = new LCS<string>(rw.readList(), rw.readList3());
            LCS<string> lcs3 = new LCS<string>(rw.readList(), rw.readList4());
            int[] count = new int[] { lcs1.iDemo(), lcs2.iDemo(), lcs3.iDemo()};
            //版本号的字符串
            string str = "MYWORD1.0";
            byte[] strArray = System.Text.Encoding.Default.GetBytes(str);
            int strlength = strArray.Length;
            //考试要求的字符串
            string require = lcs.Demo();
            byte[] reArray = System.Text.Encoding.Default.GetBytes(require);
            int relength = reArray.Length;
            //名字的字符串数组
            string path = writeMsg.Path;
            string[] sname = wm.sname(path);
            //ID的字符串数组
            string[] sid = wm.sid(path);
            //c是考生数量
            int c = writeMsg.count;

            //写版本号
            byte[] istrlength = BitConverter.GetBytes(strlength);
            F.Write(istrlength, 0, 4);
            F.Write(strArray, 0, strlength);
            //写考试要求
            byte[] irelength = BitConverter.GetBytes(relength);
            F.Write(irelength, 0, 4);
            F.Write(reArray, 0, relength);
            //写考生数量
            byte[] cbuffer = BitConverter.GetBytes(c);
            F.Write(cbuffer, 0, cbuffer.Length);
            //写考生信息
            for(int a= 0; a < sid.Length; a++)
            {
                byte[] idArray = System.Text.Encoding.Default.GetBytes(sid[a]);
                int idlength = idArray.Length;
                byte[] nameArray = System.Text.Encoding.Default.GetBytes(sname[a]);
                int namelength = nameArray.Length;
                //每个考生总字节数
                int num = idlength + namelength + count[a].ToString().Length;
                //写总数
                byte[] numbuffer = BitConverter.GetBytes(num);
                F.Write(numbuffer, 0, numbuffer.Length);
                //写id
                byte[] iidlength = BitConverter.GetBytes(idlength);
                F.Write(iidlength, 0, 4);
                F.Write(idArray, 0, idlength);
                //写name
                byte[] inamelength = BitConverter.GetBytes(namelength);
                F.Write(inamelength, 0, 4);
                F.Write(nameArray, 0, namelength);
                //写成绩
                byte[] scorebuffer = BitConverter.GetBytes(count[a]);
                F.Write(scorebuffer, 0, scorebuffer.Length);

            }
            

            //F.Position = 0;
            F.Close();
        }
    }
}
