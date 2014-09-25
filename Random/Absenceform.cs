using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Random
{
    public partial class Absenceform : Form
    {
        private ArrayList absence;
        private ArrayList absinfo = new ArrayList();
        private String filePath;
        private bool end = false;
        int i = 0;
        public bool excelpath(String path)
        {
            filePath = path;
            if (filePath != null)
                return true;
            return false;
        }
        public bool absset(ArrayList abs)
        {
            if (abs == null)
                return false;
            absence = (ArrayList)abs.Clone();
            if (absence != null)
                return true;
            return false;
        }
        public Absenceform()
        {
            InitializeComponent();
        }

        private void absence_Load(object sender, EventArgs e)
        {
            groupBox1.Visible = false;
            groupBox2.Visible = false;
            groupBox3.Visible = false;
            button1.Enabled = false;
            button2.Enabled = false;
            excelio cell = excelio.getInstance();
            if (filePath != null)
            {
                cell.openfile(filePath);
                if (absence != null)
                {
                    button2.Enabled = true;
                    foreach (int sn in absence)
                    {
                        absinfo.Add(new info(cell.getstudentname(sn), sn, cell.getabsencenum(sn)));
                    }
                    if (absinfo.Count >= 3)
                    {
                        if (absinfo.Count == 3)
                            button2.Enabled = false;
                        for (i = 0; i < 3; i++)
                        {
                            info student = (info)absinfo[i];
                            fresh(student, i);
                        }
                    }
                    else
                    {
                        button2.Enabled = false;
                        for (i = 0; i < absinfo.Count; i++)
                        {
                            info student = (info)absinfo[i];
                            fresh(student, i);
                        }
                    }
                }
            }
        }

        private void Absenceform_FormClosed(object sender, FormClosedEventArgs e)
        {
            MainForm form = (MainForm)this.Owner;
            form.Visible = true;
            this.pic1.Dispose();
            this.pic2.Dispose();
            this.pic3.Dispose();
        }
        private void fresh(info std,int i)
        {
            PictureBox pic = null;
            Label sn = null;
            Label name = null;
            Label abs = null;
            GroupBox box = null;
            if(i==0)
            {
                pic = pic1;
                sn = sn1;
                name = name1;
                abs = absnum1;
                box = groupBox1;
            }
            if (i == 1)
            {
                pic = pic2;
                sn = sn2;
                name = name2;
                abs = absnum2;
                box = groupBox2;
            }
            if (i == 2)
            {
                pic = pic3;
                sn = sn3;
                name = name3;
                abs = absnum3;
                box = groupBox3;
            }
            pic.LoadAsync(Application.StartupPath + @"/photos/" + std.getsn() + @".jpg");
            sn.Text = std.getsn().ToString();
            name.Text = std.getname();
            abs.Text = std.getabsnum().ToString();
            box.Visible = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            groupBox1.Visible = false;
            groupBox2.Visible = false;
            groupBox3.Visible = false;
            button1.Enabled = true;
            if (i >= absinfo.Count - 3)
            {
                end = true;
                button2.Enabled = false;
            }
            if(i<absinfo.Count-3)
                for (int j = 0; j < 3; j++,i++)
                {
                    info student = (info)absinfo[i];
                    fresh(student, j);
                }
                else
                    for (int j = 0; i < absinfo.Count; j++,i++)
                    {
                        info student = (info)absinfo[i];
                        fresh(student, j);
                    }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            groupBox1.Visible = false;
            groupBox2.Visible = false;
            groupBox3.Visible = false;
            button2.Enabled = true;
            if (end)
                if (i % 3 != 0)
                    i -= i % 3 + 3;
                else
                    i -= 6;
            else
                i -= 6;
            if (i == 0)
                button1.Enabled = false;
            for (int j = 0; j < 3; j++, i++)
            {
                info student = (info)absinfo[i];
                fresh(student, j);
            }
            end = false;
        }
    }
    public class info
    {
        private String name;
        private int sn;
        private int time;
        public info(String name, int sn, int time)
        {
            this.name = name;
            this.sn = sn;
            this.time = time;
        }
        public int getsn()
        {
            return sn;
        }
        public int getabsnum()
        {
            return time;
        }
        public String getname()
        {
            return name;
        }
    }
}
