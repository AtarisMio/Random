using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Cells;
using System.Collections;
using System.Windows.Forms;
using System.IO;

namespace Random
{
    
    class excelio
    {
        String filepath;
        private Workbook workbook = null;
        private Cells cells = null;
        private static excelio instance = null;
        private SortedList<long,String> namelist = new SortedList<long ,String>();
        private String classname;


        public static excelio getInstance()//单例模式
        {
            if (instance == null)
                instance = new excelio();
            return instance;
        }


        private excelio()
        {}


        public bool openfile(String filepath)//打开给定文档
        {
            this.filepath = filepath;
            workbook = new Workbook(filepath);
            if (workbook == null)
                return false;
            cells = workbook.Worksheets[0].Cells;
            if (cells == null)
                return false;
            return true;
        }


        public SortedList<long, String> readfile()//从文档中读取出所有学生姓名，并添加到namelist中
        {
            if (namelist.Count != 0)
                namelist.Clear();
            if (findString("A为缺席") == null)//已经读取过将不读取最后的解释命令
            {
                for (int i = 6; i < cells.MaxDataRow + 1; i++)
                {
                    long sn = long.Parse(cells[i, 2].StringValue.Trim());
                    String name = cells[i, 3].StringValue.Trim();
                    namelist.Add(sn, name);
                    
                }
            }
            else
            {
                for (int i = 6; i < cells.MaxDataRow; i++)
                {
                    long sn = long.Parse(cells[i, 2].StringValue.Trim());
                    String name = cells[i, 3].StringValue.Trim();
                    namelist.Add(sn, name);
                }
            }
            return namelist;
        }

        public int getabsencenum(long sn, int i = 0, int j = -1)//得到学号为sn的学生缺勤次数
        {
            int num = 0;
            int row = 0;
            row = find(sn)[0];
            for (; i <= (j == -1 ? cells.MaxDataColumn : j); i++)
            {
                if (cells[row, i + 9].StringValue == "A")
                    num++;
            }
            return num;
        }
        public ArrayList getabstime(long sn)//得到学号为sn的学生缺勤列表
        {
            ArrayList time = new ArrayList();
            int row = find(sn)[0];
            for (int i = 9; i < cells.MaxDataColumn; i++)
            {
                if (cells[row, i].StringValue == "A")
                    time.Add(gettime(i));
            }
            return time;
        }
        public ArrayList gettimes()//得到所有记录时间
        {
            ArrayList time = new ArrayList();
            for (int i = 9; i < getweek()+8; i++)
            {
                time.Add(cells[cells.MaxDataRow, i].StringValue);
            }
            return time;
        }
        private String gettime(int Column)//得到第Column列记录时间
        {
            return cells[cells.MaxDataRow, Column].StringValue;
        }
        public int getstudentnum()//取得当前课程中实际人数
        {
            if (findString("A为缺席") == null)
                return cells.MaxDataRow - 5;
            return cells.MaxDataRow - 6;
        }
        public ArrayList getabsstu(int Column, int k = 0, int j = -1)//得到第Column列中缺勤学生的列表
        {
            ArrayList result = new ArrayList();
            for (int i = 6; i < cells.MaxDataRow; i++)
            {
                if (cells[i, Column].StringValue == "A")
                {
                    student std = new student();
                    std.Serial = int.Parse(cells[i, 1].StringValue);
                    std.Sn = long.Parse(cells[i, 2].StringValue);
                    std.Name = cells[i, 3].StringValue;
                    std.Abs = getabsencenum(std.Sn, k, j);
                    result.Add(std);
                }
            }
            return result;
        }

        public String getclassname()//取得当前课程名
        {
            classname = cells[3, 11].StringValue.Trim();
            return cells[3, 11].StringValue.Trim();
        }


        public void settime()//设置第i次点名的时间记录
        {
            int i = getweek();
            int row;
            if (findString("A为缺席") == null)
            {
                row = cells.MaxDataRow + 1;
            }
            else
            {
                row = cells.MaxDataRow;
            }
            cells[row, 2].PutValue("A为缺席");
            cells[row, 8 + i].PutValue(DateTime.Now.Year + "年" + DateTime.Now.Month + "月" + DateTime.Now.Day + "日" + DateTime.Now.Hour + "时" + DateTime.Now.Minute + "分");
            
        }


        public bool save()//保存文件
        {
            try
            {
                workbook.Save(filepath);
            }
            catch
            {
                //String e = ex.ToString();
                try
                {
                    File.SetAttributes(filepath, FileAttributes.Normal);
                    workbook.Save(filepath);
                    File.SetAttributes(filepath, FileAttributes.Hidden);
                }
                catch
                {
                    MessageBox.Show("请关闭Excel后点击确认");
                    File.SetAttributes(filepath, FileAttributes.Normal);
                    workbook.Save(filepath);
                    File.SetAttributes(filepath, FileAttributes.Hidden);
                }
            }
            return true;
        }
        public bool save(String path)//保存文件(path)
        {
            try
            {
                workbook.Save(path);
            }
            catch
            {
                try
                {
                    File.SetAttributes(filepath, FileAttributes.Normal);
                    workbook.Save(filepath);
                    File.SetAttributes(filepath, FileAttributes.Hidden);
                }
                catch
                {
                    MessageBox.Show("请关闭Excel后点击确认");
                    File.SetAttributes(filepath, FileAttributes.Normal);
                    workbook.Save(filepath);
                    File.SetAttributes(filepath, FileAttributes.Hidden);
                }
            }
            return true;
        }
        public Workbook newworkbook()
        {
            return new Workbook();
        }
        public excelio newcells(Workbook w)
        {
            this.workbook = w;
            cells = workbook.Worksheets[0].Cells;
            return this;
        }
        public void sort(int key = 1)
        {
            DataSorter sorter = workbook.DataSorter;
            sorter.Order1 = Aspose.Cells.SortOrder.Ascending;
            sorter.Key1 = key;
            sorter.Key2 = 0;
            sorter.Sort(cells, 2, 0, cells.MaxDataRow, cells.MaxDataColumn);
        }
        public int[] find(long sn)//找到学号为sn的学生位置
        {
            Cell temp = cells.Find(sn.ToString(), null, new FindOptions());
            if (temp == null)
                return null;
            int[] p = new int[2];
            p[0] = temp.Row;
            p[1] = temp.Column;
            return p;
        }
        public String getstudentname(long sn)
        {
            int row = find(sn)[0];
            return cells[row, 3].StringValue;
        }

        public int[] findString(String str)//找到字符串为str的位置
        {
            Cell temp = cells.Find(str, null, new FindOptions());
            if (temp == null)
                return null;
            int[] p = new int[2];
            p[0] = temp.Row;
            p[1] = temp.Column;
            return p;
        }


        public bool mark(long sn)//标记学号为sn的学生在第i次点名缺课
        {
            int i = getweek();
            int[] position = find(sn);
            cells[position[0], 8 + i].PutValue("A");
            Style style = new Style();
            style.ForegroundColor = System.Drawing.Color.Red;
            style.Pattern = BackgroundType.Solid;
            cells[position[0], 8 + i].SetStyle(style);
            return true;
        }


        public int getweek()//取得当前为第几次点名
        {
            int lastrow;
            int[] p = findString("A为缺席");
            if (p == null)
                lastrow = cells.MaxDataRow + 1;
            else
                lastrow = cells.MaxDataRow;
            for(int i=1;i<cells.MaxDataColumn+1;i++)
            {
                if (cells[lastrow, i+8].StringValue.Trim() == "")
                    return i;
            }
            return 0;
        }
        public void setstr(String str,int row,int column)
        {
            cells[row, column].PutValue(str);
        }
    }
    public class student
    {
        private String name;

        public String Name
        {
            get { return name; }
            set { name = value; }
        }
        private int serial;

        public int Serial
        {
            get { return serial; }
            set { serial = value; }
        }
        private long sn;

        public long Sn
        {
            get { return sn; }
            set { sn = value; }
        }
        private int abs;

        public int Abs
        {
            get { return abs; }
            set { abs = value; }
        }
    }
}
