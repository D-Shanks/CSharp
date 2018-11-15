using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace test4
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //输出题目
            ReadWord rw = new ReadWord();
            LCS<string> lcs = new LCS<string>(rw.readList(),rw.readList1());
            msg.Text = lcs.ToString();
            msg.Visible = true;
            //先清除所有行
            dataGridView1.Rows.Clear();
            //初始化fileRead获取文件
            fileRead fr = new fileRead();
            writeMsg wm = new writeMsg();
            wm.road(textBox1.Text);
            fileWtite fw = new fileWtite();
            fw.writefile();
            fr.readfile();
            //扫描路径
            DirectoryInfo dir = new DirectoryInfo(textBox1.Text);

            FileSystemInfo[] files = dir.GetFileSystemInfos();
             
            int index = files.Length;
            int[] score = fileRead.score.ToArray();
            string[] name = fileRead.name.ToArray();
            string[] id = fileRead.id.ToArray();
            int i;
            for(int a = 0; a < index; a++)
            {
                i = dataGridView1.Rows.Add();
                dataGridView1.Rows[i].Cells[0].Value = id[a];
                dataGridView1.Rows[i].Cells[1].Value = name[a];
                dataGridView1.Rows[i].Cells[2].Value = score[a];
            }
        }

        private void fileselect_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";

            //在弹出窗口之前 ，设置默认文件夹
            if (textBox1.Text != null)
                dialog.SelectedPath = textBox1.Text;

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string savePath = dialog.SelectedPath;
                //把结果填写到文本框textBox1中
                textBox1.Text = savePath;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string strPath = textBox1.Text;
            ErrorMsg.Text = "";
            if (strPath == "")
            {
                ErrorMsg.Text = "目录为空";
            }
            else
            {

                DirectoryInfo directoryInfo = new DirectoryInfo(strPath);
                if (directoryInfo.Exists == false)//如果不存在
                {
                    ErrorMsg.Text = "目录不存在";
                }
            }
            if (ErrorMsg.Text != "")
            {
                ErrorMsg.Visible = true;
                button1.Enabled = false;
            }
            else
            {
                ErrorMsg.Visible = false;
                button1.Enabled = true;
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Text =
                Properties.Settings.Default.DocPath;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            //将textBox1中的文字保存到用户变量DocPath中
            Properties.Settings.Default.DocPath = textBox1.Text;
            Properties.Settings.Default.Save();
        }

    }
}
