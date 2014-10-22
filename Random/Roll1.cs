using System;
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
    public partial class Roll1 : Form
    {
        Thread thd1;
        excelio cell;
        bool start = true;
        bool flag = true;
        private int num = 1;
        private ArrayList absence = new ArrayList();
        private String filePath;
        private static String classname;
        SortedList<long, String> namelist;
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
            Roll1.classname = classname;
            if (Roll1.classname != null)
                return true;
            return false;
        }
        public Roll1()
        {
            Control.CheckForIllegalCrossThreadCalls = false; 
            InitializeComponent();
        }

        private void Roll1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (checkBox1.Checked)
            {
                absence.Add(long.Parse(label1.Text));
                cell.mark(long.Parse(label1.Text));
            }
            if(thd1!=null)
                if (thd1.IsAlive)
                    thd1.Abort();
            MainForm form1 = (MainForm)this.Owner;
            form1.Visible = true;
            form1.absset(absence);
            cell.settime();
            Random.clear();
            cell.save();
            
            this.pictureBox1.Dispose();
        }

        private void Roll1_Load(object sender, EventArgs e)
        {

            
            checkBox1.Enabled = false;
            cell = excelio.getInstance();
            //MessageBox.Show(filePath);
            cell.openfile(filePath);
            namelist = cell.readfile();
            studentnum = cell.getstudentnum();
            this.Owner.Visible = false;
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
                absence.Add(long.Parse(label1.Text));
                cell.mark(long.Parse(label1.Text));
            }
            if (flag)
            {
                if (label1.Text != "学号")
                {
                    Random instance = Random.getInstance();
                    instance.setnumber(namelist.IndexOfKey(long.Parse(label1.Text)));
                }
                checkBox1.Checked = false;
                checkBox1.Enabled = false;
                start = true;
                button1.Text = "停";
                thd1 = new Thread(new ThreadStart(lantern));
                thd1.IsBackground = true;
                thd1.Start();
                flag = false;
            }
            else
            {
                start = false;
                checkBox1.Enabled = true;
                flag = true;
                button1.Text = "点名";
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
                        while (MessageBox.Show(this, "剩余学生不足！", "提示", MessageBoxButtons.OK) != DialogResult.OK) ;
                        
                        button1.PerformClick();
                        break;
                    }
                    long sn = namelist.Keys[(int)result[0]];
                    label1.Text = sn.ToString();
                    label2.Text = namelist[sn];
                    label3.Text = "缺勤" + cell.getabsencenum(sn) + "次";
                    pictureBox1.LoadAsync(Application.StartupPath + @"/photos/" + sn + @".jpg");
                }
            }
        }     
    }
}
