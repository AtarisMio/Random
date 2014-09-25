﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace Random
{
    public partial class Roll2 : Form
    {
        Thread thd1;
        excelio cell;
        bool start = true;
        bool flag = true;
        private int num = 1;
        private ArrayList absence = new ArrayList();
        private String filePath;
        private static String classname;
        SortedList<int, String> namelist;
        private int studentnum;
        private delegate void DelegateFunction();
        public bool numset(int i)
        {
            if (i == 1 || i == 2 || i == 4)
            {
                num = i;
                return true;
            }
            return false;
        }

        //public bool absset(ArrayList abs)
        //{
        //    if (abs == null)
        //        return false;
        //    absence = (ArrayList)abs.Clone();
        //    if (absence != null)
        //        return true;
        //    return false;
        //}
        public bool excelpath(String path)
        {
            filePath = path;
            if (filePath != null)
                return true;
            return false;
        }
        public static bool classnameset(String classname)
        {
            Roll2.classname = classname;
            if (Roll2.classname != null)
                return true;
            return false;
        }
        public Roll2()
        {
            Control.CheckForIllegalCrossThreadCalls = false; 
            InitializeComponent();
        }

        private void Roll2_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (checkBox1.Checked)
            {
                absence.Add(int.Parse(label1.Text));
                cell.mark(int.Parse(label1.Text));
            }
            if (checkBox2.Checked)
            {
                absence.Add(int.Parse(label3.Text));
                cell.mark(int.Parse(label3.Text));
            }
            if (thd1 != null) if (thd1.IsAlive)
                thd1.Abort();
            MainForm form1 = (MainForm)this.Owner;
            form1.Visible = true;
            form1.absset(absence);
            cell.settime();
            Random.clear();
            cell.save();
            
            this.pictureBox1.Dispose();
            this.pictureBox2.Dispose();
        }

        private void Roll2_Load(object sender, EventArgs e)
        {
            checkBox1.Enabled = false;
            checkBox2.Enabled = false;
            cell = excelio.getInstance();
            cell.openfile(filePath);
            namelist = cell.readfile();
            studentnum = cell.getstudentnum();
            //String classname = cell.getclassname();

            //int[] position = cell.find(1131000078);
            //cell.mark(1131000078, 1);
            //int week = cell.getweek();
            //cell.settime(1);
            //cell.save();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                absence.Add(int.Parse(label1.Text));
                cell.mark(int.Parse(label1.Text));
            }
            if (checkBox2.Checked)
            {
                absence.Add(int.Parse(label3.Text));
                cell.mark(int.Parse(label3.Text));
            }
            if (flag)
            {
                if (label1.Text != "学号")
                {
                    Random instance = Random.getInstance();
                    instance.setnumber(namelist.IndexOfKey(int.Parse(label1.Text)));
                    instance.setnumber(namelist.IndexOfKey(int.Parse(label3.Text)));
                }
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox1.Enabled = false;
                checkBox2.Enabled = false;
                start = true;
                button1.Text = "停止随机";
                thd1 = new Thread(new ThreadStart(lantern));
                thd1.IsBackground = true;
                thd1.Start();
                flag = false;
            }
            else
            {
                start = false;
                checkBox1.Enabled = true;
                checkBox2.Enabled = true;
                flag = true;
                button1.Text = "开始随机";
            }

        }
        private void lantern()
        {
            Object obj = new Object();
            lock (obj)
            {
                Random instance = Random.getInstance();
                instance.rannumber(studentnum);
                while (true)
                {
                    Thread.Sleep(100);
                    if (!start)
                        break;
                    ArrayList result = instance.get(2);
                    if (result.Contains(-1))
                    {
                        MessageBox.Show("剩余学生不足！");
                        button1.PerformClick();
                        break;
                    }
                    int sn = namelist.Keys[(int)result[0]];
                    label1.Text = sn.ToString();
                    label2.Text = namelist[sn];
                    label5.Text = "缺勤" + cell.getabsencenum(sn) + "次";
                    pictureBox1.LoadAsync(Application.StartupPath + @"/photos/" + sn + @".jpg");
                    sn = namelist.Keys[(int)result[1]];
                    label3.Text = sn.ToString();
                    label4.Text = namelist[sn];
                    label6.Text = "缺勤" + cell.getabsencenum(sn) + "次";
                    pictureBox2.LoadAsync(Application.StartupPath + @"/photos/" + sn + @".jpg");
                }
            }
        }


    }
}
