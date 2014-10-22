using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Security.AccessControl;
using System.Text;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace Random
{
    public partial class MainForm : Form
    {
        TypeConverter converter = new TypeConverter();
        int num = 1;
        ArrayList absence;
        String classname = null;
        public bool numset(int i)
        {
            if (i == 1 || i == 2 || i == 4)
            {
                num = i;
                return true;
            }
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
        public bool classnameset(String classname)
        {
            this.classname = System.Convert.ToBase64String(Encoding.UTF8.GetBytes(classname));
            return true;
        }
        
        public MainForm()
        {
            InitializeComponent();
        }
        private void 打开OToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = Application.StartupPath;
            openFileDialog1.Title = "导入";
            openFileDialog1.FileName = "";
            openFileDialog1.CheckFileExists = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (openFileDialog1.FileName != Application.StartupPath + @"\" + openFileDialog1.SafeFileName)
                {
                    File.Copy(openFileDialog1.FileName, Application.StartupPath + @"\" + openFileDialog1.SafeFileName, true);
                }
                File.SetAttributes(Application.StartupPath + @"\" + openFileDialog1.SafeFileName, FileAttributes.Hidden);
                //FileInfo fi = new FileInfo(Application.StartupPath + @"\" + openFileDialog1.SafeFileName);
                //fi.Attributes = FileAttributes.Hidden;
                //fi = null;
                if (openFileDialog1.FileName != Application.StartupPath + @"\" + openFileDialog1.SafeFileName)
                {
                    ComboBoxItem cbi = new ComboBoxItem();
                    cbi.Text = openFileDialog1.SafeFileName;
                    cbi.Value = Application.StartupPath + @"\" + openFileDialog1.SafeFileName;
                    fileselect.Items.Add(cbi);
                    fileselect.SelectedIndex = 0;
                    MessageBox.Show("导入成功");
                }
                else
                {
                    MessageBox.Show("您选择的文件为正在工作的文件，不能导入");
                }
                button1.Enabled = true;
                
            }
            //DirectoryInfo d = new DirectoryInfo(Application.StartupPath);
            //SortedList<String, String> FileList = GetCurrentDirAllxls(d);
            //foreach (KeyValuePair<string, string> item in FileList)
            //{
            //    ComboBoxItem cbi = new ComboBoxItem();
            //    cbi.Text = item.Key;
            //    cbi.Value = item.Value;
            //    fileselect.Items.Add(cbi);
            //}
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            ConfigSectionData data = config.Sections["add"] as ConfigSectionData;
            if (data != null)
                num = data.Randomnum;
            SortedList<String, String> FileList = new SortedList<String, String>();
            //读取配置文件
            String snum = ConfigurationManager.AppSettings["randomnum"];
            if (snum != null && snum != "")
                num = int.Parse(snum);
            //DirectoryInfo d = new DirectoryInfo(Application.StartupPath);
            FileList = GetCurrentDirAllxls(Application.StartupPath, FileList);
            foreach (KeyValuePair<string, string> item in FileList)
            {
                ComboBoxItem cbi = new ComboBoxItem();
                cbi.Text = item.Key;
                cbi.Value = item.Value;
                fileselect.Items.Add(cbi);
            }
            if(fileselect.Items.Count>0)
                fileselect.SelectedIndex = 0;
            config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            ConfigurationSectionGroup group1 = config.SectionGroups["A" + Math.Abs(fileselect.SelectedText.GetHashCode()).ToString()];
            if (group1 != null)
            {
                data = config.SectionGroups["A" + Math.Abs(fileselect.SelectedText.GetHashCode()).ToString()].Sections["add"] as ConfigSectionData;
                if (data.Absence.Length>0)
                    absence = (ArrayList)RetrieveObject(System.Convert.FromBase64String(data.Absence));
                if (classname != null)
                    classname = data.Classname;
            }
            if (fileselect.Items.Count == 0)
                button1.Enabled = false;
        }
        private SortedList<String, String> GetCurrentDirAllxls(String dir,SortedList<String, String> FileList)
        {
            String[] allFile = Directory.GetFileSystemEntries(dir, "*.xls");
            foreach (String fi in allFile)
            {
                if (!fi.EndsWith("_点名记录.xls"))
                    FileList.Add(fi.Substring(fi.LastIndexOf('\\') + 1), fi);
            }
            return FileList;
        }

        private void 退出XToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void 选项OToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Settings form = new Settings();
            form.Owner = this;
            form.numset(num);//设置配置文件中的人数
            form.Show();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            ConfigSectionData data = new ConfigSectionData();
            data.Randomnum = num;
            ConfigurationSection add = config.Sections["add"] as ConfigurationSection;
            if (add == null)
                config.Sections.Add("add", data);
            else
            {
                config.Sections.Remove("add");
                config.Sections.Add("add", data);
            }
            
            data = new ConfigSectionData();
            if(absence!=null)
                data.Absence = System.Convert.ToBase64String(GetBinaryFormatData(absence));
            if(classname!=null)
                data.Classname = classname;
            ConfigurationSectionGroup group1 = config.SectionGroups["A" + Math.Abs(fileselect.SelectedText.GetHashCode()).ToString()];
            if (group1 == null)
                config.SectionGroups.Add("A" + Math.Abs(fileselect.SelectedText.GetHashCode()).ToString(), new ConfigurationSectionGroup());
            group1 = config.SectionGroups["A" + Math.Abs(fileselect.SelectedText.GetHashCode()).ToString()];
            add = group1.Sections["add"] as ConfigurationSection;
            if (add == null)
                config.SectionGroups["A" + Math.Abs(fileselect.SelectedText.GetHashCode()).ToString()].Sections.Add("add", data);
            else
            {
                //System.Collections.Specialized.NameObjectCollectionBase.KeysCollection a = group1.Sections.Keys;
                group1.Sections.Remove("add");
                
                group1.Sections.Add("add", data);
            }
            config.Save(ConfigurationSaveMode.Full);

        }

        private void fileselect_SelectionChangeCommitted(object sender, EventArgs e)
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            ConfigurationSectionGroup group1 = config.SectionGroups["A" + Math.Abs(fileselect.SelectedText.GetHashCode()).ToString()];
            if (group1 != null)
            {
                ConfigSectionData data = config.SectionGroups["A" + Math.Abs(fileselect.SelectedText.GetHashCode()).ToString()].Sections["add"] as ConfigSectionData;
                if (data.Absence.Length>0)
                    absence = (ArrayList)RetrieveObject(System.Convert.FromBase64String(data.Absence));
                if (classname != null)
                    classname = data.Classname;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            excelio cell = excelio.getInstance();
            if(absence!=null)
                absence.Clear();
            String filepath = (fileselect.SelectedItem as ComboBoxItem).Value.ToString();
            if (classname != null)
            {
                cell.openfile(filepath);
                filepath = filepath.Substring(0, filepath.LastIndexOf("\\") + 1) + Encoding.UTF8.GetString(System.Convert.FromBase64String(classname));
            }
            if (num == 1)
            {
                Roll1 form = new Roll1();
                //form.absset(absence);
                form.excelpath((fileselect.SelectedItem as ComboBoxItem).Value.ToString());
                form.Owner = this;
                form.Show();
            }
            if (num == 2)
            {
                Roll2 form = new Roll2();
                //form.absset(absence);
                form.excelpath((fileselect.SelectedItem as ComboBoxItem).Value.ToString());
                form.Owner = this;
                form.Show();
            }
            if (num == 4)
            {
                Roll4 form = new Roll4();
                //form.absset(absence);
                form.excelpath((fileselect.SelectedItem as ComboBoxItem).Value.ToString());
                form.Owner = this;
                form.Show();
            }
        }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2202:不要多次释放对象")]
        public byte[] GetBinaryFormatData(object dsOriginal)
        {
            byte[] binaryDataResult = null;
            MemoryStream memStream = new MemoryStream();
            BinaryFormatter brFormatter = new BinaryFormatter();
            brFormatter.Serialize(memStream, dsOriginal);
            binaryDataResult = memStream.ToArray();
            memStream.Close();
            //memStream.Dispose();
            return binaryDataResult;
        }
        public object RetrieveObject(byte[] binaryData)
        {
            MemoryStream memStream = new MemoryStream(binaryData);
            BinaryFormatter brFormatter = new BinaryFormatter();
            Object obj = brFormatter.Deserialize(memStream);
            return obj;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            excelio cell = excelio.getInstance();
            if (absence == null)
                absence = new ArrayList();
            Absenceform form = new Absenceform();
            form.Owner = this;
            form.absset(absence);
            if (fileselect.Items.Count != 0)
            {
                cell.openfile((fileselect.SelectedItem as ComboBoxItem).Value.ToString());
                form.excelpath((fileselect.SelectedItem as ComboBoxItem).Value.ToString());
            }
            this.Visible = false;
            form.Show();
        }

        private void 打印PToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (fileselect.SelectedItem != null)
            {
                Export form = new Export();
                form.Owner = this;
                this.Visible = false;
                form.excelpath((fileselect.SelectedItem as ComboBoxItem).Value.ToString());
                form.Show();
            }
            else
                MessageBox.Show("未选择要导出的文件");
        }

        private void 关于AToolStripMenuItem_Click(object sender, EventArgs e)
        {
            about form = new about();
            form.Show();
        }



    }
    class ConfigSectionData : ConfigurationSection
    {
        [ConfigurationProperty("randomnum")]
        public int Randomnum
        {
            get { return (int)this["randomnum"]; }
            set { this["randomnum"] = value; }
        }

        [ConfigurationProperty("absence")]
        public String Absence
        {
            get { return (String)this["absence"]; }
            set { this["absence"] = value.Clone(); }
        }
        public String Classname
        {
            get { return (String)this["classname"]; }
            set { this["classname"] = value; }
        }
    }
}
