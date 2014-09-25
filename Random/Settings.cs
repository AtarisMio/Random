using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Random
{
    public partial class Settings : Form
    {
        private int num = 1;
        public bool numset(int i)
        {
            if(i == 1||i==2||i==4)
            {
                num = i;
                return true;
            }
            return false;
        }
        public Settings()
        {
            InitializeComponent();
        }


        private void Settings_Load(object sender, EventArgs e)
        {
            if (num == 1)
                radioButton1.Checked = true;
            if (num == 2)
                radioButton2.Checked = true;
            if (num == 4)
                radioButton3.Checked = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
                num = 1;
            if (radioButton2.Checked)
                num = 2;
            if (radioButton3.Checked)
                num = 4;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Settings_FormClosed(object sender, FormClosedEventArgs e)
        {
            MainForm form1 = (MainForm)this.Owner;
            form1.numset(num);
            form1.Visible = true;
        }
    }
}
