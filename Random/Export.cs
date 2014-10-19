using Aspose.Cells;
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
    public partial class Export : Form
    {
        int list = 0;
        excelio cell;
        private ArrayList checkdate;
        private String filePath;
        private String classname;
        public bool excelpath(String path)
        {
            filePath = path;
            if (filePath != null)
                return true;
            return false;
        }
        public Export()
        {
            InitializeComponent();
        }

        private void Export_Load(object sender, EventArgs e)
        {
            cell = excelio.getInstance();
            cell.openfile(filePath);
            checkdate = cell.gettimes();
            classname = cell.getclassname();
            foreach (String item in checkdate)
            {
                ComboBoxItem cbi = new ComboBoxItem();
                cbi.Text = item;
                cbi.Value = list;
                dateselect.Items.Add(cbi);
                list++;
            }
        }

        private void dateselect_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (dateselect.CheckedItems.Count == list)
                checkBox1.Checked = true;
            if (dateselect.CheckedItems.Count < list)
                checkBox1.Checked = false;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (!checkBox1.Checked)
                dateselect.ClearSelected();
            else
                for (int j = 0; j < dateselect.Items.Count; j++)
                    dateselect.SetItemChecked(j, true);  
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ArrayList abssttdbydate = new ArrayList();
            int first = dateselect.CheckedIndices[0];
            int last = dateselect.CheckedIndices.Count - 1;
            foreach(ComboBoxItem date in dateselect.CheckedItems)
            {
                ArrayList absstdbysn = new ArrayList();
                foreach(student st in cell.getabsstu((int)date.Value+9,first,last))
                {
                    absstdbysn.Add(st);
                }
                abssttdbydate.Add(absstdbysn);
            }//将旧表信息提取
            Workbook workbook = cell.newworkbook();
            cell = cell.newcells(workbook);
            int j = 1, k = 3;
            cell.setstr("序号", 1, 0);
            cell.setstr("学号", 1, 1);
            cell.setstr("姓名", 1, 2);
            cell.setstr("缺勤次数", 1, 3);
            foreach(ArrayList absstdbysn in abssttdbydate)
            {
                
                cell.setstr("第" + (k - 2) + "次", 0, k+1);
                foreach(student st in absstdbysn)
                {
                    if (cell.find(st.Sn) != null)
                    {
                        
                        int temp = j;
                        j = cell.find(st.Sn)[0];
                        cell.setstr(st.Serial.ToString(), j, 0);
                        cell.setstr(st.Sn.ToString(), j, 1);
                        cell.setstr(st.Name, j, 2);
                        cell.setstr("缺勤", j, k+1);
                        j = temp;
                    }
                    else
                    {
                        j++;
                        cell.setstr(st.Abs.ToString(), j, 3);
                        cell.setstr(st.Serial.ToString(), j, 0);
                        cell.setstr(st.Sn.ToString(), j, 1);
                        cell.setstr(st.Name, j, 2);
                        cell.setstr("缺勤", j, k+1);
                    }
                }
                k++;
            }
            k = 0;
            foreach (ComboBoxItem date in dateselect.CheckedItems)
            {
                cell.setstr(date.Text, 1, k+4);
                k++;
            }
            cell.sort(3);
            cell.save(Application.StartupPath+"\\"+classname+"点名记录.xls");
            MessageBox.Show("导出成功");
            this.Close();
        }
        //private void button1_Click(object sender, EventArgs e)
        //{
        //    Workbook workbook = cell.newworkbook();
        //    cell = cell.newcells(workbook);
        //    int j = 1, k = 3;
        //    ArrayList abssttdbytime = new ArrayList();
        //    cell.setstr("序号", 0, 0);
        //    cell.setstr("学号", 0, 0);
        //    cell.setstr("姓名", 0, 0);
        //    foreach (ComboBoxItem date in dateselect.CheckedItems)
        //    {
        //        cell.setstr("第" + (k - 2) + "次", 0, k);
        //        foreach (student st in cell.getabsstu((int)date.Value + 8))
        //        {
        //            if (cell.find(st.Sn) != null)
        //            {
        //                int temp = j;
        //                j = cell.find(st.Sn)[0];
        //                cell.setstr(st.Serial.ToString(), j, 0);
        //                cell.setstr(st.Sn.ToString(), j, 1);
        //                cell.setstr(st.Name, j, 2);
        //                cell.setstr("缺勤", j, k);
        //                j = temp;
        //            }
        //            else
        //            {
        //                j++;
        //                cell.setstr(st.Serial.ToString(), j, 0);
        //                cell.setstr(st.Sn.ToString(), j, 1);
        //                cell.setstr(st.Name, j, 2);
        //                cell.setstr("缺勤", j, k);
        //            }
        //        }
        //        k++;
        //    }
        //    for (int i = 0; i < dateselect.CheckedItems.Count; i++)
        //    {
        //        ComboBoxItem date = (ComboBoxItem)dateselect.SelectedItems[k];
        //        cell.setstr(date.Text, j, i + 2);
        //    }
        //    cell.save(Application.StartupPath + "\\" + classname + "点名记录.xls");

        //}
        private void Export_FormClosed(object sender, FormClosedEventArgs e)
        {
            MainForm form = (MainForm)this.Owner;
            form.Visible = true;
        }

        
    }
    
}
