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
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ArrayList abssttdbydate = new ArrayList();
            if (dateselect.CheckedItems.Count != 0)
            {
                int first = dateselect.CheckedIndices[0];
                int last = dateselect.CheckedIndices.IndexOf(dateselect.CheckedIndices.Count-1);
                if((last-first)>=dateselect.CheckedItems.Count)
                {
                    MessageBox.Show("暂只支持连续输出");
                    return;
                }
                foreach (ComboBoxItem date in dateselect.CheckedItems)
                {
                    ArrayList absstdbysn = new ArrayList();
                    foreach (student st in cell.getabsstu((int)date.Value + 9, first, last))
                    {
                        absstdbysn.Add(st);
                    }
                    if (absstdbysn.Count != 0)
                    {
                        abssttdbydate.Add(new KeyValuePair<ArrayList, KeyValuePair<String, int>>(absstdbysn, new KeyValuePair<String, int>(date.Text, dateselect.Items.IndexOf(date))));
                    }
                }//将旧表信息提取
                Workbook workbook = cell.newworkbook();
                cell = cell.newcells(workbook);
                int j = 1, k = 3;
                cell.setstr("序号", 1, 0);
                cell.setstr("学号", 1, 1);
                cell.setstr("姓名", 1, 2);
                cell.setstr("缺勤次数", 1, 3);
                foreach (KeyValuePair<ArrayList, KeyValuePair<String, int>> absstdbysns in abssttdbydate)
                {
                    ArrayList absstdbysn = absstdbysns.Key;
                    KeyValuePair<String, int> dates = absstdbysns.Value;
                    cell.setstr("第" + (dates.Value + 1) + "次", 0, k + 1);
                    cell.setstr(dates.Key, 1, k + 1);
                    foreach (student st in absstdbysn)
                    {
                        if (cell.find(st.Sn) != null)
                        {

                            int temp = j;
                            j = cell.find(st.Sn)[0];
                            cell.setstr(st.Serial.ToString(), j, 0);
                            cell.setstr(st.Sn.ToString(), j, 1);
                            cell.setstr(st.Name, j, 2);
                            cell.setstr("缺勤", j, k + 1);
                            j = temp;
                        }
                        else
                        {
                            j++;
                            cell.setstr(st.Abs.ToString(), j, 3);
                            cell.setstr(st.Serial.ToString(), j, 0);
                            cell.setstr(st.Sn.ToString(), j, 1);
                            cell.setstr(st.Name, j, 2);
                            cell.setstr("缺勤", j, k + 1);
                        }
                    }
                    k++;
                }
                k = 0;
                /*foreach (ComboBoxItem date in dateselect.CheckedItems)
                {
                    cell.setstr(date.Text, 1, k+4);
                    k++;
                }*/
                cell.sort(3);
                cell.save(Application.StartupPath + "\\" + classname + "_点名记录.xls");
                MessageBox.Show("导出成功");
                cell.openfile(filePath);
                this.Close();
            }
            else
            {
                MessageBox.Show("未选则输出日期");
            }
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

        private void checkBox1_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
                for (int j = 0; j < dateselect.Items.Count; j++)
                    dateselect.SetItemChecked(j, true);
            else
                for (int j = 0; j < dateselect.Items.Count; j++)
                    dateselect.SetItemChecked(j, false);
        }

        private void Export_MouseUp(object sender, MouseEventArgs e)
        {
            if (dateselect.CheckedItems.Count == list)
                checkBox1.CheckState = CheckState.Checked;
            else
            {
                if (dateselect.CheckedItems.Count == 0)
                    checkBox1.CheckState = CheckState.Unchecked;
                else
                {
                    checkBox1.CheckState = CheckState.Indeterminate;
                }
            }
        }

        private void dateselect_KeyUp(object sender, KeyEventArgs e)
        {
            if (dateselect.CheckedItems.Count == list)
                checkBox1.CheckState = CheckState.Checked;
            else
            {
                if (dateselect.CheckedItems.Count == 0)
                    checkBox1.CheckState = CheckState.Unchecked;
                else
                {
                    checkBox1.CheckState = CheckState.Indeterminate;
                }
            }
        }

        
    }
    
}
