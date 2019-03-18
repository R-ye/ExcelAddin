using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace ExcelAddIn2
{
    public partial class Form3 : Form
    {
        static Microsoft.Office.Interop.Excel._Workbook wbook = Globals.ThisAddIn.Application.ActiveWorkbook;//获取激活的工作簿
        Worksheet Worksheet = wbook.Worksheets[1];//获取名为sheet1的工作表


        public Form3()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
        }
        public bool IsFloat(string str)
        {
            try
            {
                string regextext = @"^(-?\d+)(\.\d+)?$";
                Regex regex = new Regex(regextext, RegexOptions.None);
                return regex.IsMatch(str.Trim());
            }
            catch
            {
                return false;
            }
        }
        //判断字符串是否为整数
        public bool IsInteger(string str)
        {
            try
            {
                int i = Convert.ToInt32(str);
                return true;
            }
            catch
            {
                return false;
            }
        }
        public bool IsCount(string str)
        {
            try
            {
                string pattern = @"^[0-9]*$";

                if (Regex.IsMatch(str.Trim(), pattern))
                {
                    if (str.Length == 19 && str.Substring(0, 1) == "6")
                    {
                        return true;
                    }
                    if (str.Length == 23 && str.Substring(0, 1) == "1")
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            catch
            {
                return false;
            }
        }
        public bool IsOther(string str)
        {
            string pattern = @"^[0-9]*$";

            if (Regex.IsMatch(str.Trim(), pattern))
            {if (str == "0" || str == "0.0" || str == "0.00")
                {
                    return true;
                }
                if ((str.Length >= 9 && str.Length != 23 && str.Length != 19) )
                {
                    return true;
                }

                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }
        public string[,] RemoveColNull(string[,] Rdata)
        {
            List<int> rrow = new List<int>();
            int row1 = Rdata.GetLength(0);
            int col1 = Rdata.GetLength(1);
            for (int i = 0; i < row1; i++)
            {
                //progressBar1.Value = i * 100 / (row1 - 1);
                int jk = 0;
                for (int j = 0; j < col1; j++)
                {
                   
                    if (Rdata[i, j].Replace("\n", "").Replace(" ", "").Replace("\t", "").Replace("\r", "") == "")
                    {
                        jk++;
                    }

                }
                if (jk == Rdata.GetLength(1))
                {
                    rrow.Add(i);//满足一行均为空的就记录下来这个行
                }
            }
            string[,] ret = new string[row1 - rrow.Count, col1];
            int p = 0;
            for (int i = 0; i < row1; i++)
            { if (rrow.Contains(i) == false)
                {
                    for (int j = 0; j < col1; j++)
                    {
                        ret[p, j] = Rdata[i, j];
                    }
                    p++;
                }
            }
            return ret;

        }
        public string[,] RemoveRowNull(string[,] Rdata)
        {
            List<int> rrow = new List<int>();
            int row1 = Rdata.GetLength(0);
            int col1 = Rdata.GetLength(1);
            for (int i = 0; i < col1; i++)//列
            {
                //progressBar1.Value = i * 100 / (col1 - 1);
                int jk = 0;
                for (int j = 0; j < row1; j++)
                {
                    if (Rdata[j, i].Replace("\n", "").Replace(" ", "").Replace("\t", "").Replace("\r", "") == "")
                    {
                        jk++;
                    }

                }
                if (jk == Rdata.GetLength(0))
                {
                    rrow.Add(i);//满足一行均为空的就记录下来这个行
                }
            }
            string[,] ret = new string[row1, col1 - rrow.Count];
            int p = 0;
            int rt = 0;
            for (int i = 0; i < col1; i++)
            {
                if (rrow.Contains(i) == false)
                {
                    for (int j = 0; j < row1; j++)
                    {
                        ret[j, p] = Rdata[j, i];
                    }
                    p++;
                }
                else
                {
                    rt++;
                }
                
            }
            if (rt != 0)
            {
                this.listView1.Items.Add(new ListViewItem(new string[] { "数据行处理反馈", "无效列共计" + rt + "列", "!?", "0", "[不返回数据]" }));
                listView1.Items[this.listView1.Items.Count - 1].UseItemStyleForSubItems = true;
                listView1.Items[this.listView1.Items.Count - 1].BackColor = Color.Green;
                listView1.Items[this.listView1.Items.Count - 1].ForeColor = Color.White;
            }
            return ret;

        }
        public string[,] IsHaveCount(string[,] Rdata)//对有账号的进行保留
        {
            try
            {
                List<int> rrow = new List<int>();
                int row1 = Rdata.GetLength(0);
                int col1 = Rdata.GetLength(1);
                for (int i = 0; i < row1; i++)//行
                {
                    progressBar1.Value = i * 100 / (row1 - 1);
                    bool HaveCount = false;
                    for (int j = 0; j < col1; j++)
                    {
                        if (IsCount(Rdata[i, j]))
                        {
                            HaveCount = true;
                        }
                    }
                    if (HaveCount == false)
                    {
                        rrow.Add(i);
                    }
                }
                // t = rrow.Count;
                string[,] ret = new string[row1 - rrow.Count, col1];
                int p = 0;
                int rt = 0;
                for (int i = 0; i < row1; i++)
                {
                    if (rrow.Contains(i) == false)
                    {
                        for (int j = 0; j < col1; j++)
                        {
                            ret[p, j] = Rdata[i, j].Replace("\n", "").Replace(" ", "").Replace("\t", "").Replace("\r", "");

                        }
                        p++;
                    }
                    else
                    {
                        rt++;
                    }
                }
                if (rt != 0)
                {
                    this.listView1.Items.Add(new ListViewItem(new string[] { "数据账号处理反馈", "无效账号共计" + rt + "组", "!?", "0", "[不返回数据]" }));
                    listView1.Items[this.listView1.Items.Count - 1].UseItemStyleForSubItems = true;
                    listView1.Items[this.listView1.Items.Count - 1].BackColor = Color.Green;
                    listView1.Items[this.listView1.Items.Count - 1].ForeColor = Color.White;
                }
                return ret;
            }
            catch(Exception e)
            {
               // MessageBox.Show(e.ToString());
                return null;
            }
        }
        public string[,] IsHaveName(string[,] Rdata)//对有账号的进行保留
        {
            List<int> rrow = new List<int>();
            int row1 = Rdata.GetLength(0);
            int col1 = Rdata.GetLength(1);
            for (int i = 0; i < row1; i++)//行
            {
              //  progressBar1.Value = i * 100 / (row1 - 1);
                bool HaveCount = false;
                for (int j = 0; j < col1; j++)
                {
                    if (IsName(Rdata[i, j].Replace("\n", "").Replace(" ", "").Replace("\t", "").Replace("\r", "")))
                    {
                        HaveCount = true;
                    }
                }
                if (HaveCount == false)
                {
                    rrow.Add(i);
                }
            }
            // t = rrow.Count;
            string[,] ret = new string[row1 - rrow.Count, col1];
            int p = 0;
            int rt = 0;
            for (int i = 0; i < row1; i++)
            {
                if (rrow.Contains(i) == false)
                {
                    for (int j = 0; j < col1; j++)
                    {
                        ret[p, j] = Rdata[i, j].Replace("\n", "").Replace(" ", "").Replace("\t", "").Replace("\r", "");
                    }
                    p++;
                }
                else
                {
                    rt++;
                }
            }
            if (rt != 0)
            {
                this.listView1.Items.Add(new ListViewItem(new string[] { "数据姓名处理反馈", "无效姓名共计" + rt + "组", "!?", "0", "[不返回数据]" }));
                listView1.Items[this.listView1.Items.Count - 1].UseItemStyleForSubItems = true;
                listView1.Items[this.listView1.Items.Count - 1].BackColor = Color.Green;
                listView1.Items[this.listView1.Items.Count - 1].ForeColor = Color.White;
            }
            return ret;
        }
        public string[,] IsOtherNum(string[,] Rdata)//对有账号的进行保留
        {
            List<int> rrow = new List<int>();
            int row1 = Rdata.GetLength(0);
            int col1 = Rdata.GetLength(1);
            for (int i = 0; i < col1; i++)//列
            {
             //   progressBar1.Value = i * 100 / (col1 - 1);
                int jk = 0;
                for (int j = 0; j < row1; j++)
                {
                    if (IsOther(Rdata[j, i].Replace("\n", "").Replace(" ", "").Replace("\t", "").Replace("\r", "").Replace("X", "").Replace("x", "")))
                    {
                        jk++;
                    }

                }
                if (jk == Rdata.GetLength(0))
                {
                    rrow.Add(i);//满足一行均为空的就记录下来这个行
                }
            }
            string[,] ret = new string[row1, col1 - rrow.Count];
            int p = 0;
            for (int i = 0; i < col1; i++)
            {
                if (rrow.Contains(i) == false)
                {
                    for (int j = 0; j < row1; j++)
                    {
                        ret[j, p] = Rdata[j, i];
                    }
                    p++;
                }
            }
            return ret;

        }
        public string[,] IsOtherChar(string[,] Rdata)//对重复出现的删除
        {
            List<int> rrow = new List<int>();
            int row1 = Rdata.GetLength(0);
            int col1 = Rdata.GetLength(1);
            for (int i = 0; i < col1; i++)//列
            {
             //   progressBar1.Value = i * 100 / (col1 - 1);
                int jk = 0;
                for (int j = 0; j < row1; j++)
                {
                    if ((Rdata[j, i].IndexOf("月") != -1 && Rdata[j,i].Length>=4) || Rdata[j, i].IndexOf("村") != -1 || Rdata[j, i].IndexOf("组") != -1 || Rdata[j, i].IndexOf("队") != -1 || Rdata[j, i].IndexOf("街") != -1 || Rdata[j, i].Length <= 1 || Rdata[j, i].IndexOf("山") != -1 || Rdata[j, i].IndexOf("乡") != -1)
                    {
                        jk++;
                    }

                }
                if (jk >= Rdata.GetLength(0) / 2)
                {
                    rrow.Add(i);//满足一行均为空的就记录下来这个行
                }
            }
            string[,] ret = new string[row1, col1 - rrow.Count];
            int p = 0;
            int rt = 0;
            for (int i = 0; i < col1; i++)
            {
                if (rrow.Contains(i) == false)
                {
                    for (int j = 0; j < row1; j++)
                    {
                        ret[j, p] = Rdata[j, i];
                    }
                    p++;
                }
                else
                {
                    rt++;
                }
            }
            if (rt != 0)
            {
                this.listView1.Items.Add(new ListViewItem(new string[] { "数据特殊项处理反馈", "无效项目共计" + rt + "组", "!?", "0", "[不返回数据]" }));
                listView1.Items[this.listView1.Items.Count - 1].UseItemStyleForSubItems = true;
                listView1.Items[this.listView1.Items.Count - 1].BackColor = Color.Green;
                listView1.Items[this.listView1.Items.Count - 1].ForeColor = Color.White;
            }
            return ret;

        }
        public bool IsName(string text)//判断是否是名字
        { if (System.Text.RegularExpressions.Regex.IsMatch(text, @"[\u4e00-\u9fbb]+$") && text.Length >= 2 && text.Length <= 4)
            {

                return true;
            }
            else
            { return false; }
            // return System.Text.RegularExpressions.Regex.IsMatch(text, @"[\u4e00-\u9fbb]+$");
        }
        public bool IsMoney1(string str)//判断是否有小数
        {
            Regex reg = new Regex(@"^\d+\.\d+$");
            if (reg.IsMatch(str))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public  double  GetSimilarityWith(string sourceString, string str)
        {

            double Kq = 2;
            double Kr = 1;
            double Ks = 1;

            char[] ss = sourceString.ToCharArray();
            char[] st = str.ToCharArray();

            //获取交集数量
            int q = ss.Intersect(st).Count();
            int s = ss.Length -q;
            int r = st.Length- q;

            return Kq * q / (Kq * q + Kr * r + Ks * s);
        }
        public string[,] HaveSameCol(string[,] Rdata,ref int a)
        {
            StringCompute same = new StringCompute();
            List<int> coll = new List<int>();
            int row1 = Rdata.GetLength(0);
            int col1 = Rdata.GetLength(1);
            for (int i = 0; i < col1; i++)//列
            {
          progressBar1.Value = i * 100 / (col1 - 1);
                {
                    for (int p = (i + 1); p < col1; p++)
                    {
                        string temp = "";
                        string temp1 = "";
                            for (int j = 0; j < row1; j++)//行
                            {
                                temp = temp + Rdata[j, i];//基准数据
                                temp1 = temp1 + Rdata[j, p];//比较数据
                            }
                       same.SpeedyCompute(temp,temp1);    // 计算相似度， 不记录比较时间
                        decimal rate = same.ComputeResult.Rate;         // 相似度百分之几，完全匹配相似度为1
                       // listBox1.Items.Add("列"+i+"--列"+p+"相似度："+ rate);
                       if (rate>=(decimal)0.60 && coll.Contains(p)==false)
                       {
                          coll.Add(p);
                        }
                    }

                }
            }
            a = coll.Count;
            string[,] ret = new string[row1, col1 - coll.Count];
            int p1 = 0;
            int rt = 0;
            for (int i = 0; i < col1; i++)
            {
                if (coll.Contains(i) == false)
                {
                    for (int j = 0; j < row1; j++)
                    {
                        ret[j, p1] = Rdata[j, i];
                    }
                    p1++;
                }
                else
                {
                    rt++;
                }
            }
            if (rt != 0)
            {
                this.listView1.Items.Add(new ListViewItem(new string[] { "数据重复项处理反馈", "重复项共计" + rt + "组", "!?", "0", "[不返回数据]" }));
                listView1.Items[this.listView1.Items.Count - 1].UseItemStyleForSubItems = true;
                listView1.Items[this.listView1.Items.Count - 1].BackColor = Color.Green;
                listView1.Items[this.listView1.Items.Count - 1].ForeColor = Color.White;
            }
            return ret;
        }
        public string[,] HaveXuhao(string[,] Rdata)
        {
            StringCompute same = new StringCompute();
            List<int> coll = new List<int>();
            int row1 = Rdata.GetLength(0);
            int col1 = Rdata.GetLength(1);
            for (int i = 0; i < col1; i++)//列
            {

              //  progressBar1.Value = i * 100 / (col1 - 1);

                string temp = "";
                        string temp1 = "";
               

                        for (int j = 0; j < row1; j++)//行
                        {
                            temp = temp + Rdata[j, i];//基准数据
                             temp1 = temp1 + j;
                   
                        }
                        same.SpeedyCompute(temp, temp1);    // 计算相似度， 不记录比较时间
                        decimal rate = same.ComputeResult.Rate;         // 相似度百分之几，完全匹配相似度为1
                                                                        //listBox1.Items.Add("列"+i+"--列"+p+"相似度："+ rate);
                        if ((rate >= (decimal)0.55 && coll.Contains(i) == false) )
                        {
                            coll.Add(i);
                        }
                   
            }
           
            string[,] ret = new string[row1, col1 - coll.Count];
            int p1 = 0;
            int rt = 0;
            for (int i = 0; i < col1; i++)
            {
                if (coll.Contains(i) == false)
                {
                    for (int j = 0; j < row1; j++)
                    {
                        ret[j, p1] = Rdata[j, i];
                    }
                    p1++;
                }
                else
                {
                    rt++;
                }
            }
            if (rt != 0)
            {
                this.listView1.Items.Add(new ListViewItem(new string[] { "数据序号处理反馈", "序号项共计" + rt + "组，已自动剔除", "!?", "0", "[不返回数据]" }));
                listView1.Items[this.listView1.Items.Count - 1].UseItemStyleForSubItems = true;
                listView1.Items[this.listView1.Items.Count - 1].BackColor = Color.Green;
                listView1.Items[this.listView1.Items.Count - 1].ForeColor = Color.White;
            }
            return ret;
        }
        public void Init(object str)
        {
            try
            {
                toolStripStatusLabel1.Text = "数据组数：未统计";
                toolStripStatusLabel2.Text = "总金额：未统计";
                toolStripStatusLabel3.Text = "有问题数据：未统计";
                listView1.Items.Clear();
                int row, col;
                row = Worksheet.UsedRange.Rows.Count;//获取行
                col = Worksheet.UsedRange.Columns.Count;//获取列
                if (row >= 5000)
                {
                    row = 5000;
                }
                if (col >= 20)
                {
                    col = 20;
                }
                if (row <= 1 || col <= 1)
                {
                    this.listView1.Items.Add(new ListViewItem(new string[] { "置入数据", "加载失败,数据容量过小，不符合实际", "×", "-1", "[null]" }));
                    button1.Enabled = true;
                    return;
                }
                string[,] Celldata = new string[row, col];//将数据置入数组
                                                          //  linkLabel1.Text = "正在置入数据...["+row+"*"+col+"]";
                this.listView1.Items.Add(new ListViewItem(new string[] { "置入数据", "已加载", "√", "1", "[" + row + " * " + col + "]" }));
                this.listView1.EnsureVisible(this.listView1.Items.Count - 1);
                listView1.Items[this.listView1.Items.Count - 1].UseItemStyleForSubItems = true;

                for (int i = 0; i < row; i++)
                {
                    progressBar1.Value = i * 100 / (row - 1);
                    for (int j = 0; j < col; j++)
                    {
                        Celldata[i, j] = (Worksheet.Cells[i + 1, j + 1]).Text.ToString().Replace("\n", "").Replace(" ", "").Replace("\t", "").Replace("\r", "").Replace("X", "").Replace("x", "");
                    }

                }

                this.listView1.Items.Add(new ListViewItem(new string[] { "置入数据", "加载成功", "√", "2", "[" + Celldata.GetLength(0) + " * " + Celldata.GetLength(1) + "]" }));
                //  label1.Image = imageList1.Images[1];
                //  label1.Text ="    "+(DateTime.Now.ToShortTimeString() + " 置入数据完成，CAP:[" + Celldata.GetLength(0) + "*" + Celldata.GetLength(1) + "]");
                this.listView1.Items.Add(new ListViewItem(new string[] { "数据行对齐", "已加载", "√", "3", "[" + Celldata.GetLength(0) + " * " + Celldata.GetLength(1) + "]" }));
                this.listView1.EnsureVisible(this.listView1.Items.Count - 1);
                // label1.Visible = true;
                string[,] ret = RemoveColNull(Celldata);
                if (ret.Length == 0)
                {
                    this.listView1.Items.Add(new ListViewItem(new string[] { "数据无效", "加载失败", "×", "0", "[" + ret.GetLength(0) + " * " + ret.GetLength(1) + "]" }));
                    listView1.Items[this.listView1.Items.Count - 1].UseItemStyleForSubItems = true;
                    listView1.Items[this.listView1.Items.Count - 1].BackColor = Color.Red;
                    listView1.Items[this.listView1.Items.Count - 1].ForeColor = Color.White;
                    button1.Enabled = true;
                    button1.Text = "智能分析";
                    return;
                }
                this.listView1.Items.Add(new ListViewItem(new string[] { "数据行对齐完成", "加载成功", "√", "4", "[" + ret.GetLength(0) + " * " + ret.GetLength(1) + "]" }));
                this.listView1.EnsureVisible(this.listView1.Items.Count - 1);
                // linkLabel1.Text = "数据行对齐完成，CAP:[" + ret.GetLength(0) + "*" + ret.GetLength(1) + "]";
                // label2.Visible = true;
                //label2.Image = imageList1.Images[1];
                //label2.Text="    "+(DateTime.Now.ToShortTimeString() + " 数据行对齐完成，CAP:[" + ret.GetLength(0) + "*" + ret.GetLength(1) + "]");
                //RemoveRowNull
                this.listView1.Items.Add(new ListViewItem(new string[] { "数据列对齐", "已加载", "√", "4", "[" + ret.GetLength(0) + " * " + ret.GetLength(1) + "]" }));

                //   listBox1.SelectedIndex = listBox1.Items.Count - 1;
                string[,] ret1 = RemoveRowNull(ret);
                if (ret1.Length == 0)
                {
                    this.listView1.Items.Add(new ListViewItem(new string[] { "数据无效", "缺少有效数据", "×", "0", "[" + ret1.GetLength(0) + " * " + ret1.GetLength(1) + "]" }));
                    listView1.Items[this.listView1.Items.Count - 1].UseItemStyleForSubItems = true;
                    listView1.Items[this.listView1.Items.Count - 1].BackColor = Color.Red;
                    listView1.Items[this.listView1.Items.Count - 1].ForeColor = Color.White;
                    //  listBox1.SelectedIndex = listBox1.Items.Count - 1;
                    //  listBox1.SelectedIndex = listBox1.Items.Count - 1;
                    button1.Enabled = true;
                    button1.Text = "智能分析";
                    return;
                }
                this.listView1.Items.Add(new ListViewItem(new string[] { "数据列对齐完成", "加载成功", "√", "5", "[" + ret1.GetLength(0) + " * " + ret1.GetLength(1) + "]" }));
                this.listView1.EnsureVisible(this.listView1.Items.Count - 1);
                //  label3.Image = imageList1.Images[1];
                // label3.Text="    "+(DateTime.Now.ToShortTimeString() + " 数据列对齐完成，CAP:[" + ret1.GetLength(0) + "*" + ret1.GetLength(1) + "]");
                this.listView1.Items.Add(new ListViewItem(new string[] { "数据账号校验", "已加载", "√", "6", "[" + ret1.GetLength(0) + " * " + ret1.GetLength(1) + "]" }));


                string[,] ret2 = IsHaveCount(ret1);
                if (ret2.Length == 0)
                {
                    this.listView1.Items.Add(new ListViewItem(new string[] { "数据账号校验", "缺少账号", "×", "6", "[" + ret2.GetLength(0) + " * " + ret2.GetLength(1) + "]" }));
                    listView1.Items[this.listView1.Items.Count - 1].UseItemStyleForSubItems = true;
                    listView1.Items[this.listView1.Items.Count - 1].BackColor = Color.Red;
                    listView1.Items[this.listView1.Items.Count - 1].ForeColor = Color.White;

                    button1.Enabled = true;
                    button1.Text = "智能分析";
                    return;
                }
                this.listView1.Items.Add(new ListViewItem(new string[] { "数据账号校验", "加载成功", "√", "7", "[" + ret2.GetLength(0) + " * " + ret2.GetLength(1) + "]" }));
                this.listView1.EnsureVisible(this.listView1.Items.Count - 1);
                // label4.Visible = true;
                //  label4.Image = imageList1.Images[1];
                //  label4.Text="    "+(DateTime.Now.ToShortTimeString() + " 数据账号提取完成，CAP:[" + ret2.GetLength(0) + "*" + ret2.GetLength(1) + "]");
                this.listView1.Items.Add(new ListViewItem(new string[] { "数据姓名校验", "已加载", "√", "8", "[" + ret2.GetLength(0) + " * " + ret2.GetLength(1) + "]" }));


                string[,] ret3 = IsHaveName(ret2);
                if (ret3.Length == 0)
                {
                    this.listView1.Items.Add(new ListViewItem(new string[] { "数据姓名校验", "缺少姓名", "×", "0", "[" + ret3.GetLength(0) + " * " + ret3.GetLength(1) + "]" }));
                    listView1.Items[this.listView1.Items.Count - 1].UseItemStyleForSubItems = true;
                    listView1.Items[this.listView1.Items.Count - 1].BackColor = Color.Red;
                    listView1.Items[this.listView1.Items.Count - 1].ForeColor = Color.White;

                    button1.Enabled = true;
                    button1.Text = "智能分析";
                    return;
                }
                this.listView1.Items.Add(new ListViewItem(new string[] { "数据姓名校验", "加载成功", "√", "9", "[" + ret3.GetLength(0) + " * " + ret3.GetLength(1) + "]" }));
                this.listView1.EnsureVisible(this.listView1.Items.Count - 1);
                // label5.Visible = true;
                // label5.Image = imageList1.Images[1];
                // label5.Text="    "+(DateTime.Now.ToShortTimeString() + " 数据姓名提取完成，CAP:[" + ret3.GetLength(0) + "*" + ret3.GetLength(1) + "]");

                if (ret3.GetLength(1) == 2)
                {
                    this.listView1.Items.Add(new ListViewItem(new string[] { "数据姓名校验", "数据无效", "×", "-1", "[" + ret3.GetLength(0) + " * " + ret3.GetLength(1) + "]" }));

                    this.listView1.EnsureVisible(this.listView1.Items.Count - 1);
                    listView1.Items[this.listView1.Items.Count - 1].UseItemStyleForSubItems = true;
                    listView1.Items[this.listView1.Items.Count - 1].BackColor = Color.Red;
                    listView1.Items[this.listView1.Items.Count - 1].ForeColor = Color.White;
                    button1.Enabled = true;
                    button1.Text = "智能分析";
                    return;
                }
                if (ret3.GetLength(1) >= 4)
                {
                    this.listView1.Items.Add(new ListViewItem(new string[] { "数据二次处理完成", "存在序号或多列数字或多个姓名", "?", "9", "[" + ret3.GetLength(0) + " * " + ret3.GetLength(1) + "]" }));
                    this.listView1.EnsureVisible(this.listView1.Items.Count - 1);
                    listView1.Items[this.listView1.Items.Count - 1].UseItemStyleForSubItems = true;
                    listView1.Items[this.listView1.Items.Count - 1].BackColor = Color.Blue;
                    listView1.Items[this.listView1.Items.Count - 1].ForeColor = Color.White;
                    // button1.Enabled = true;
                    //  button1.Text = "智能分析";
                }
                string[,] ret4 = IsOtherNum(ret3);
                //   label6.Visible = true;
                // label6.Image = imageList1.Images[1];
                this.listView1.Items.Add(new ListViewItem(new string[] { "数据默认过滤", "已加载", "√", "9", "[" + ret4.GetLength(0) + " * " + ret4.GetLength(1) + "]" }));

                //  listBox1.SelectedIndex = listBox1.Items.Count - 1;
                if (ret4.GetLength(1) == 2)
                {
                    this.listView1.Items.Add(new ListViewItem(new string[] { "数据默认过滤", "数据无效", "×", "-1", "[" + ret4.GetLength(0) + " * " + ret4.GetLength(1) + "]" }));

                    this.listView1.EnsureVisible(this.listView1.Items.Count - 1);
                    listView1.Items[this.listView1.Items.Count - 1].UseItemStyleForSubItems = true;
                    listView1.Items[this.listView1.Items.Count - 1].BackColor = Color.Red;
                    listView1.Items[this.listView1.Items.Count - 1].ForeColor = Color.White;
                    button1.Enabled = true;
                    button1.Text = "智能分析";
                    return;
                }
                string[,] ret5 = IsOtherChar(ret4);
                // label7.Visible = true;
                // label7.Image = imageList1.Images[1];
                this.listView1.Items.Add(new ListViewItem(new string[] { "数据杂项判定", "已加载", "√", "10", "[" + ret5.GetLength(0) + " * " + ret5.GetLength(1) + "]" }));


                if (ret5.GetLength(1) == 2)
                {
                    this.listView1.Items.Add(new ListViewItem(new string[] { "数据杂项判定", "存在程序过滤关键词，且超过50%", "×", "-1", "[" + ret5.GetLength(0) + " * " + ret5.GetLength(1) + "]" }));
                    this.listView1.EnsureVisible(this.listView1.Items.Count - 1);
                    listView1.Items[this.listView1.Items.Count - 1].UseItemStyleForSubItems = true;
                    listView1.Items[this.listView1.Items.Count - 1].BackColor = Color.Red;
                    listView1.Items[this.listView1.Items.Count - 1].ForeColor = Color.White;

                    button1.Enabled = true;
                    button1.Text = "智能分析";
                    return;
                }
                //HaveSameCol
                int a = 0;
                string[,] ret6 = HaveSameCol(ret5, ref a);
                // MessageBox.Show(a.ToString());
                // label8.Visible = true;
                //  label8.Image = imageList1.Images[1];
                //  label8.Text="    "+(DateTime.Now.ToShortTimeString() + " 数据相似度分析完成，CAP:[" + ret6.GetLength(0) + "*" + ret6.GetLength(1) + "]");
                this.listView1.Items.Add(new ListViewItem(new string[] { "数据相似度分析", "已加载", "√", "11", "[" + ret6.GetLength(0) + " * " + ret6.GetLength(1) + "]" }));

                if (ret6.GetLength(1) == 2)
                {
                    this.listView1.Items.Add(new ListViewItem(new string[] { "数据相似度分析", "数据无效", "×", "11", "[" + ret6.GetLength(0) + " * " + ret6.GetLength(1) + "]" }));
                    this.listView1.EnsureVisible(this.listView1.Items.Count - 1);
                    listView1.Items[this.listView1.Items.Count - 1].UseItemStyleForSubItems = true;
                    listView1.Items[this.listView1.Items.Count - 1].BackColor = Color.Red;
                    listView1.Items[this.listView1.Items.Count - 1].ForeColor = Color.White;
                    button1.Enabled = true;
                    button1.Text = "智能分析";
                    return;
                }
                string[,] ret7 = HaveXuhao(ret6);
                // MessageBox.Show(a.ToString());
                // label9.Visible = true;
                // label9.Image = imageList1.Images[1];
                // label9.Text="    "+(DateTime.Now.ToShortTimeString() + " 数据建模完成，CAP:[" + ret7.GetLength(0) + "*" + ret7.GetLength(1) + "]");
                this.listView1.Items.Add(new ListViewItem(new string[] { "数据序号剔除", "已加载", "√", "12", "[" + ret7.GetLength(0) + " * " + ret7.GetLength(1) + "]" }));

                if (ret7.GetLength(1) == 2)
                {
                    this.listView1.Items.Add(new ListViewItem(new string[] { "数据序号剔除", "数据无效", "×", "-1", "[" + ret7.GetLength(0) + " * " + ret7.GetLength(1) + "]" }));
                    this.listView1.EnsureVisible(this.listView1.Items.Count - 1);
                    linkLabel1.Text = "数据无效";
                    listView1.Items[this.listView1.Items.Count - 1].UseItemStyleForSubItems = true;
                    listView1.Items[this.listView1.Items.Count - 1].BackColor = Color.Red;
                    listView1.Items[this.listView1.Items.Count - 1].ForeColor = Color.White;
                    button1.Enabled = true;
                    button1.Text = "智能分析";
                    return;
                }
                this.listView1.Items.Add(new ListViewItem(new string[] { "数据交换", "已加载", "√", "13", "[" + ret7.GetLength(0) + " * " + ret7.GetLength(1) + "]" }));
                this.listView1.EnsureVisible(this.listView1.Items.Count - 1);
                //对数据进行分类排列，第一列是序号，第二列账号，第三列是姓名，第四列是金额。只需要取Ret的[0,0][0,1][0,2]。。。
                List<int> zh = new List<int>();
                List<int> zf = new List<int>();
                List<int> je = new List<int>();
                for (int i = 0; i < ret7.GetLength(1); i++)
                {
                    if (IsCount(ret7[0, i]))
                    {
                        zh.Add(i);
                        continue;
                    }
                    if (IsName(ret7[0, i]))
                    {
                        zf.Add(i);
                        continue;
                    }
                    if (IsFloat(ret7[0, i]) || (IsInteger(ret7[0, i]) && (ret7[0, i]).Length <= 6))
                    {
                        je.Add(i);
                        continue;
                    }
                }
                string[,] ret8 = new string[ret7.GetLength(0), ret7.GetLength(1)];
                double sum = 0.0;
                int u = 0;
                for (int i = 0; i < ret8.GetLength(0); i++)
                {
                    for (int j = 0; j < zh.Count; j++)
                    {
                        ret8[i, 0 + j] = ret7[i, zh[j]];
                        //  op = j+1;
                    }
                    for (int j = 0; j < zf.Count; j++)
                    {
                        ret8[i, 1 + j] = ret7[i, zf[j]];
                        //if(j==zf.Count-1)
                        // op = op+j;
                    }
                    for (int j = 0; j < je.Count; j++)
                    {
                        ret8[i, 2 + j] = ret7[i, je[j]];
                        // if (j == je.Count - 1)
                        //  op = op + j;
                        if (ret7[i, je[j]] == "")
                        {
                            this.listView1.Items.Add(new ListViewItem(new string[] { "数据交换", "加载异常,缺少金额", "×?", "在第" + je[j] + "行", "[" + ret8.GetLength(0) + " * " + ret8.GetLength(1) + "]" }));
                            this.listView1.EnsureVisible(this.listView1.Items.Count - 1);
                            listView1.Items[this.listView1.Items.Count - 1].UseItemStyleForSubItems = true;
                            listView1.Items[this.listView1.Items.Count - 1].BackColor = Color.Red;
                            listView1.Items[this.listView1.Items.Count - 1].ForeColor = Color.White;
                        }
                        u++;
                        if (ret7[i, je[j]].Length != 2)
                            sum += double.Parse(0 + ret7[i, je[j]]);
                    }
                }
                if (u > ret8.GetLength(0))
                {
                    this.listView1.Items.Add(new ListViewItem(new string[] { "数据统计", "数据总金额统计有误，存在其他非金额数字", "×", "在第4列或第5列", "[" + ret8.GetLength(0) + " * " + ret8.GetLength(1) + "]" }));
                    this.listView1.EnsureVisible(this.listView1.Items.Count - 1);

                    listView1.Items[this.listView1.Items.Count - 1].UseItemStyleForSubItems = true;
                    listView1.Items[this.listView1.Items.Count - 1].BackColor = Color.Red;
                    listView1.Items[this.listView1.Items.Count - 1].ForeColor = Color.White;
                    toolStripStatusLabel2.Text = "总金额：" + sum + "[不可信]";
                }
                else
                {
                    toolStripStatusLabel2.Text = "总金额：" + sum + "[可信]";
                }
                toolStripStatusLabel1.Text = "数据组数：" + ret8.GetLength(0);

                toolStripStatusLabel3.Text = "有问题数据：" + (ret2.GetLength(0) - ret8.GetLength(0));
                Microsoft.Office.Interop.Excel.Worksheet newWorkbook = Globals.ThisAddIn.Application.Worksheets.Add(System.Type.Missing, Worksheet);
                //  Microsoft.Office.Interop.Excel.Worksheet newWorkbook = Globals.ThisAddIn.Application.Worksheets.Add(System.Type.Missing);
                Microsoft.Office.Interop.Excel.Worksheet newWorksheet1 = Globals.ThisAddIn.Application.ActiveSheet;
                newWorksheet1.Cells.NumberFormat = "@";
                long yy = 0;
                this.listView1.Items.Add(new ListViewItem(new string[] { "数据布道", "已加载", "√", "13", "[" + ret8.GetLength(0) + " * " + ret8.GetLength(1) + "]" }));
                this.listView1.EnsureVisible(this.listView1.Items.Count - 1);
                for (int i = 1; i <= ret8.GetLength(0); i++)
                {
                    progressBar1.Value = int.Parse((i * 100 / (ret8.GetLength(0))).ToString());
                    for (int j = 0; j <= ret8.GetLength(1); j++)
                    {
                        if (j == 0)
                        {
                            newWorksheet1.Cells[i, j + 1] = i;
                            newWorksheet1.Cells[i, j + 1].EntireColumn.AutoFit();

                        }
                        else
                        {
                            if (ret8[i - 1, j - 1] == "")
                            {
                                yy++;
                                newWorksheet1.Cells[i, j + 1].Interior.Color = Color.Red;//设置颜色
                                newWorksheet1.Cells[i, j + 1].Font.Color = Color.White;

                            }

                            {
                                newWorksheet1.Cells[i, j + 1] = ret8[i - 1, j - 1];
                                newWorksheet1.Cells[i, j + 1].EntireColumn.AutoFit();
                            }
                        }
                    }
                }
                this.listView1.Items.Add(new ListViewItem(new string[] { "数据布道", "数据分析结束", "√", "1", "[" + ret8.GetLength(0) + " * " + ret8.GetLength(1) + "]" }));
                this.listView1.EnsureVisible(this.listView1.Items.Count - 1);
                // label10.Visible = true;
                //  label10.Image = imageList1.Images[1];
                if (yy != 0)
                {
                    this.listView1.Items.Add(new ListViewItem(new string[] { "数据布道", "存在数据缺失" + yy + "项", "×", "-1", "[" + ret8.GetLength(0) + " * " + ret8.GetLength(1) + "]" }));
                    this.listView1.EnsureVisible(this.listView1.Items.Count - 1);
                    toolStripStatusLabel3.Text = "有问题数据：>=" + yy;
                    listView1.Items[this.listView1.Items.Count - 1].UseItemStyleForSubItems = true;
                    listView1.Items[this.listView1.Items.Count - 1].BackColor = Color.Red;
                    listView1.Items[this.listView1.Items.Count - 1].ForeColor = Color.White;
                }
                else
                {

                    this.listView1.Items.Add(new ListViewItem(new string[] { "数据布道", "布道完成", "√", "1", "[" + ret8.GetLength(0) + " * " + ret8.GetLength(1) + "]" }));
                    this.listView1.EnsureVisible(this.listView1.Items.Count - 1);
                }
                button1.Enabled = true;
                button1.Text = "智能分析";
            }
            catch(Exception e)
            {
                MessageBox.Show("出现致命错误，请反馈给开发者，联系Email:admin@congjingstudio.cn\r\n错误编码："+e.Source+"\r\n"+e.StackTrace+"\r\n"+e.Message);
            }
            //this.Close();
            //ret1中列和行都有数据，但可能为空
        }
        /*  public void Init(object str)
          {
              listBox1.Items.Clear();
              int row, col;
              int xuhao = 0;
              string error = "";
              int je = 0;
              int xm = 0;
              int zh = 0;
              string name = Worksheet.Name;
              linkLabel1.Text = "正在获取数据源...";
              listBox1.Items.Add"    "+(DateTime.Now.ToShortTimeString()+">>工作表名 "+name);
              listBox1.SelectedIndex = listBox1.Items.Count - 1;
              Thread.Sleep(100);
              linkLabel1.Text = "正在对数据源进行属性分析...";
              Thread.Sleep(100);
              row = Worksheet.UsedRange.Rows.Count;
              col = Worksheet.UsedRange.Columns.Count;
           //   listBox1.Items.Add"    "+(DateTime.Now.ToShortTimeString() + ">> 表属性：" + Worksheet.UsedRange.Rows.Count+" * " +Worksheet.UsedRange.Columns.Count);
              if (row == 1 || col== 1)
              { listBox1.Items.Add"    "+(DateTime.Now.ToShortTimeString() + ">> "+name + " 表不满足分析条件，自动停止分析！");
                  listBox1.SelectedIndex = listBox1.Items.Count - 1;
                  return;
              }
              else
              {
                  listBox1.Items.Add"    "+(DateTime.Now.ToShortTimeString() + ">> "+name + " 表满足分析条件");
                  listBox1.SelectedIndex = listBox1.Items.Count - 1;
              }
              Thread.Sleep(100);
              linkLabel1.Text = "数据预处理中...";
              List<int> rows = new List<int>();
              for (int i = 1; i <= row; i++)
              {
                  linkLabel1.Text = "自动过滤数据中 "+int.Parse((i*100/row).ToString()) +"%";
                  int temp1 = 0, temp2 = 0, temp3 = 0;
                  for (int j = 1; j <= col; j++) 
                  {

                      {
                          string data = ((Microsoft.Office.Interop.Excel.Range)Worksheet.Cells[i, j]).Text.ToString() + "#";
                          //listBox1.Items.Add(data);
                          if (data == "#")
                          {
                              continue;
                          }
                          if (data.Replace("#", "").Length >= 1 && data.Replace("#", "").Length <= 10 && (IsFloat(data.Replace("#", "")) || IsInteger(data.Replace("#", ""))))
                          {
                              temp1 = temp1 + 1;//表示为金额或者序号
                              je = j;
                              continue;
                          }
                          if (data.Replace("#", "").Length == 19 || data.Replace("#", "").Length == 23)
                          {
                              temp2 = temp2 + 1;//表示为账号
                              zh = j;
                              continue;
                          }
                          if (CheckStringChineseReg(data.Replace("#", "")) && (data.Replace("#", "").Length >= 2 && data.Replace("#", "").Length <= 4))
                          {
                              temp3 = temp3 + 1;
                              xm = j;
                              continue;
                          }


                      }
                      }
                  if (temp1 + temp2 + temp3 < 3 || temp2 == 0 || temp3 == 0 || temp1 == 0)
                  {
                      for (int j = 1; j <= col; j++)
                      {
                          //Font.OutlineFont = True
                          Range data1 = ((Range)(Worksheet.Cells[i, j]));
                          data1.Font.Strikethrough = true;
                          // (Range)(Worksheet.Cells[i, j])).c
                        data1.Interior.Color= Color.Red;//设置颜色
                          data1.Font.Color = Color.White;
                      }
                  }
                  else
                  {
                      rows.Add(i);//加入有效数据中
                  }
              }
              linkLabel1.Text = "自动过滤完成！";
              if (rows.Count == 0)
              {
                  listBox1.Items.Add"    "+(DateTime.Now.ToShortTimeString() + ">> 无有效数据，请检查");
                  listBox1.SelectedIndex = listBox1.Items.Count - 1;
                  button1.Enabled = true;
                  button1.Text = "自动分析";
                  return;

              }
              listBox1.Items.Add"    "+(DateTime.Now.ToShortTimeString() + ">> 过滤无效数据,并标记为红色");
              Thread.Sleep(100);
              linkLabel1.Text = "数据布道中...";
              listBox1.Items.Add"    "+(DateTime.Now.ToShortTimeString() + ">> 有效数据 " + rows.Count + " 行");
              listBox1.SelectedIndex = listBox1.Items.Count - 1;
              listBox1.Items.Add"    "+(DateTime.Now.ToShortTimeString() + ">> 正在新表中填充有效数据");
              listBox1.SelectedIndex = listBox1.Items.Count - 1;
              Microsoft.Office.Interop.Excel.Worksheet newWorkbook = Globals.ThisAddIn.Application.Worksheets.Add(System.Type.Missing, Worksheet);
              //  Microsoft.Office.Interop.Excel.Worksheet newWorkbook = Globals.ThisAddIn.Application.Worksheets.Add(System.Type.Missing);
              Microsoft.Office.Interop.Excel.Worksheet newWorksheet1 = Globals.ThisAddIn.Application.ActiveSheet;
              int m = 0,l=1;
              foreach (int k in rows)
              {


                  m = 1;
                      for (int j = 1; j <= col; j++)
                      {

                          {
                              string data = ((Microsoft.Office.Interop.Excel.Range)Worksheet.Cells[k, j]).Text.ToString() + "#";

                              if (data == "#")
                              {

                                  continue;
                              }
                              if (data.Replace("#", "").Length >= 1 && data.Replace("#", "").Length <= 10 && (IsFloat(data.Replace("#", "")) || IsInteger(data.Replace("#", ""))))
                              {

                               newWorksheet1.Cells[l, m] = data.Replace("#", "");
                              m ++;
                              continue;
                              }
                              if (data.Replace("#", "").Length == 19 || data.Replace("#", "").Length == 23)
                              {
                              newWorksheet1.Cells[l, m] = "'"+data.Replace("#", "");
                              m++;
                              continue;
                              }
                              if (CheckStringChineseReg(data.Replace("#", "")) )
                              {
                              newWorksheet1.Cells[l, m] = data.Replace("#", "");
                              m++;
                              continue;
                              }


                          }
                      }
                  l++;
                  }
              linkLabel1.Text = "数据列交换中...";
              listBox1.Items.Add"    "+(DateTime.Now.ToShortTimeString() + ">> 数据列交换处理");
              listBox1.SelectedIndex = listBox1.Items.Count - 1;
              for (int u = 1; u <= newWorksheet1.UsedRange.Rows.Count; u++)
              {
                  string temp = "",temp1="";
                  temp = ((Microsoft.Office.Interop.Excel.Range)newWorksheet1.Cells[u, 3]).Text.ToString();//第三列提取出来
                  temp1 = ((Microsoft.Office.Interop.Excel.Range)newWorksheet1.Cells[u, 2]).Text.ToString();//第三列提取出来
                  for (int p = 1; p < newWorksheet1.UsedRange.Columns.Count; p++)
                  {
                      string data2 = (newWorksheet1.Cells[u, p]).Text.ToString();//第三列提取出来
                      if(CheckStringChineseReg(data2.Replace("#", "")))
                      {
                          newWorksheet1.Cells[u, p] = temp;//交换姓名
                          newWorksheet1.Cells[u, 3] = data2;
                        //  m++;
                          continue;
                      }
                      if (data2.Replace("#", "").Length == 19 || data2.Replace("#", "").Length == 23)
                      {
                          newWorksheet1.Cells[u, p] = temp1;//交换账号
                                                            // m++;
                          newWorksheet1.Cells[u, 2] = "'"+data2;
                          continue;
                      }
                  }
              }
              linkLabel1.Text = "数据列交换结束";
              listBox1.Items.Add"    "+(DateTime.Now.ToShortTimeString() + ">> 数据列交换成功，进行数据建模");
              button1.Enabled = true;
              button1.Text = "自动分析";
          }
  */


        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            button1.Text = "正在智能分析中...";
            Thread thread1 = new Thread(new ParameterizedThreadStart(Init));
            thread1.Start();
            

        }

        private void Form3_Load(object sender, EventArgs e)
        {

            for (int i=1;i<=wbook.Worksheets.Count;i++)
            {
                comboBox1.Items.Add(wbook.Worksheets[i].Name);
            }
            comboBox1.Text = comboBox1.Items[0].ToString();
            //listBox1.Items.Add"    "+(DateTime.Now.ToShortTimeString() + ">> 分析引擎初始化成功");
            //listBox1.SelectedIndex = listBox1.Items.Count - 1;
        }

        private void button6_Click(object sender, EventArgs e)
        {
           
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {if (comboBox1.Text != "")
            {
                Worksheet = wbook.Worksheets[comboBox1.Text];
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog file = new SaveFileDialog();//定义新的文件保存位置控件
            file.Filter = "文本文件(*.TXT)|*.TXT";//设置文件后缀的过滤
            if (file.ShowDialog() == DialogResult.OK)//如果有文件保存路径
            {
                Encoding bm = Encoding.Default;
                if (radioButton1.Checked == true)
                {
                    bm = Encoding.Default;
                }
                else
                {
                    bm = Encoding.UTF8;
                }
                StreamWriter sw = new StreamWriter(file.FileName, false, bm);
                Microsoft.Office.Interop.Excel.Worksheet newWorksheet1 = Globals.ThisAddIn.Application.ActiveSheet;
                int row = newWorksheet1.UsedRange.Rows.Count;
                int col = newWorksheet1.UsedRange.Columns.Count;
                for (int i = 1; i <= row; i++)
                {
                    string ret = "";

                    for (int j = 1; j <= col; j++)
                    {if (checkBox7.Checked == false)
                        { ret = ret + newWorksheet1.Cells[i, j].Text.ToString() ; }
                        else
                        { ret = ret + newWorksheet1.Cells[i, j].Text.ToString() + (textBox1.Text=="" ? "|":textBox1.Text); }
                    }
                    if (i != row)
                    {
                        ret = ret + "\r\n";
                    }
                    sw.Write(ret);  //写入文件中
                }
                sw.Flush();//清理缓冲区
                sw.Close();//关闭文件
                MessageBox.Show("导出成功！");
            }
        }
        private void toolStripStatusLabel6_Click(object sender, EventArgs e)
        {
            MessageBox.Show("本程序是基于数学模型进行提取，具体原理如下：\r\n1、对表格数据进行归一对齐处理，剔除掉无效数据\r\n2、对行和列进行细致分析，大致确定数据类型\r\n3、建立数据类型对应识别模型，分别对不同类型进行校验\r\n4、对金额数据和序号数据进行逆序数计算，剔除序号\r\n5、对可能重复数据计算相似度，相似度过高则仅保留一列\r\n6、对卡号进行校验、对存折账号进行位数校验，对账号中其他无关字符进行过滤\r\n7、对姓名进行分析，以可能性确定姓名，并剔除姓名中其他非法字符，例如空格、回车\r\n8、对分析后的数据进行二次处理，并对处理后的数据进行列识别和列交换\r\n9、对识别处理后数据进行生成统计继续分析，避免错误。");
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            textBox1.Enabled = checkBox7.Checked;
        }
    }
}
