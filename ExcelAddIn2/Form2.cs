using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ExcelAddIn2
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        private string DaXie(string money)
        {
            try
            {
                string s = double.Parse(money).ToString("#L#E#D#C#K#E#D#C#J#E#D#C#I#E#D#C#H#E#D#C#G#E#D#C#F#E#D#C#.0B0A");
                string d = Regex.Replace(s, @"((?<=-|^)[^1-9]*)|((?'z'0)[0A-E]*((?=[1-9])|(?'-z'(?=[F-L\.]|$))))|((?'b'[F-L])(?'z'0)[0A-L]*((?=[1-9])|(?'-z'(?=[\.]|$))))", "${b}${z}");
                return Regex.Replace(d, ".", delegate (Match m) { return "负元空零壹贰叁肆伍陆柒捌玖空空空空空空空分角拾佰仟万億兆京垓秭穰"[m.Value[0] - '-'].ToString(); });
            }
            catch
            {
                return "";
            }
        }
        public void ADD(ComboBox c,string path)
        {

            if (!c.Items.Contains(c.Text) && c.Text !="")
            {
               
            

            //文件追加 true

            StreamWriter sw = new StreamWriter(path, true);

            if (!(c.Items.Contains(c.Text)))
            {

                sw.WriteLine(c.Text);

            }


            sw.Close();
                //保存combobox的选项内容到配置文件1.ini
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;

            ADD(comboBox1,"khh.ini");

            ADD(comboBox2, "fkdw.ini");
            ADD(comboBox3, "zh.ini");

            ADD(comboBox4, "skh.ini");

            ADD(comboBox5, "skmc.ini");
            ADD(comboBox6, "yt.ini");

            Write("name.txt",textBox4.Text);






            //Microsoft.Office.Interop.Excel._Workbook wbook = Globals.ThisAddIn.Application.ActiveWorkbook;//获取激活的工作簿
            //  group1.Label = "当前共有 "+wbook.Sheets.Count.ToString()+" 张表\r\n自动引导至第一张表，表名为："+wbook.Sheets[1].Name;//获取第一个工作表;

            Microsoft.Office.Interop.Excel.Worksheet newWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            newWorksheet.Cells.NumberFormat = "@";
            // newWorksheet.Name = "Sheet1";
            //  Worksheet worksheet = wbook.Worksheets[1];//获取名为sheet1的工作表

            //  Microsoft.Office.Interop.Excel.Worksheet newWorkbook = Globals.ThisAddIn.Application.Worksheets.Add(System.Type.Missing, worksheet);
            SetWidth("A", 6.5);
            SetWidth("B", 16.88);
            SetWidth("C", 11);
            SetWidth("D", 22.63);
            SetWidth("E", 10.75);
            SetWidth("F", 20.25);
            SetMerg("A1", "F1");
            SetMerg("A6", "C6");
            SetMerg("A5", "B5");
            SetMerg("D6", "F6");
            SetMerg("D8", "E8");
            SetMerg("E9", "F9");
            SetFont("A1", "宋体", 16, "安徽青阳农商行" + textBox2.Text+ "大额资金支付审批表", true, true, 34.5);
            SetFont("E2", "仿宋_GB2312", 12, "划款日期", false, true, 23.25);
            SetFont("F2", "仿宋_GB2312", 12, textBox2.Text, false, true, 23.25);
            SetFont("A3", "仿宋_GB2312", 12, "开户行", false, true, 50.25 );
            SetFont("C3", "仿宋_GB2312", 12, "付款单位", false, true, 50.25 );
            SetFont("E3", "仿宋_GB2312", 12, "账号", false, true, 50.25);
            SetFont("B3", "仿宋_GB2312", 12, comboBox1.Text, false, true, 50.25,true);
            SetFont("D3", "仿宋_GB2312", 12, comboBox2.Text, false, true, 50.25,true);
            SetFont("F3", "仿宋_GB2312", 12, "'"+comboBox3.Text, false, true, 50.25,true);
            SetFont("A4", "仿宋_GB2312", 12, "收款行", false, true, 50.25);
            SetFont("B4", "仿宋_GB2312", 12, comboBox4.Text, false, true, 50.25);
            SetFont("C4", "仿宋_GB2312", 12, "收款人名称", false, true, 50.25);
            SetFont("D4", "仿宋_GB2312", 12, comboBox5.Text, false, true, 50.25,true);
            SetFont("E4", "仿宋_GB2312", 12, "用途", false, true, 50.25);
            SetFont("F4", "仿宋_GB2312", 12, comboBox6.Text, false, true, 50.25);
            SetFont("A5", "仿宋_GB2312", 12, "金额（万元）", false, true, 48.75);
            SetFont("C5", "仿宋_GB2312", 12, "大写", false, true, 48.75,true);
            SetFont("E5", "仿宋_GB2312", 12, "小写", false, true, 48.75);
            SetFont("A6", "仿宋_GB2312", 12, "支行审批意见：", false, true, 67.50);
            SetFont("D8", "仿宋_GB2312", 14, "支行行长或副行长签字：", false, true, 26.25);
            SetFont("D5", "仿宋_GB2312", 12, textBox3.Text, false, true, 48.75,true);
            SetFont("F5", "仿宋_GB2312", 12, textBox1.Text, false, true, 48.75,true);
            SetFont("E9", "仿宋_GB2312", 14, textBox2.Text, false, true, 18.75,true);
            Microsoft.Office.Interop.Excel.Range range = newWorksheet.get_Range("A3", "F6");
            //   range.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternCrissCross;
            // range.Borders.Weight = 1;
            range.Cells.Borders.LineStyle = 1;
            // range.Borders.get_Item(XlBordersIndex.xlDiagonalDown).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            //  range.Borders.get_Item(XlBordersIndex.xlDiagonalUp).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            //  range.Borders.get_Item(XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            //  range.Borders.get_Item(XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            //  range.Borders.get_Item(XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDot;
            button1.Enabled = true;
            newWorksheet.PageSetup.TopMargin = 1.3f;
            newWorksheet.PageSetup.BottomMargin = 1.3f;
            newWorksheet.PageSetup.LeftMargin = 1.2f;
            newWorksheet.PageSetup.RightMargin = 0.4f;
            newWorksheet.PageSetup.CenterHorizontally = true;
           // MessageBox.Show(Globals.ThisAddIn.Application.ActiveWorkbook.FullName);
            this.Close();
        }

        public void SetFont(string a, string font, int size, string content, bool b, bool iscenter, double height,bool iswarp=false )
        {
            Microsoft.Office.Interop.Excel.Worksheet newWorksheet1 = Globals.ThisAddIn.Application.ActiveSheet;
            newWorksheet1.get_Range(a, Missing.Value).Font.Name = font;
            newWorksheet1.get_Range(a, Missing.Value).Font.Bold = b;
            newWorksheet1.get_Range(a, Missing.Value).Font.Size = size;
            newWorksheet1.get_Range(a, Missing.Value).Value2 = content;
            if (iscenter == true)
            {
                newWorksheet1.get_Range(a, Missing.Value).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            }
            newWorksheet1.Rows[a.Substring(1, 1), System.Type.Missing].RowHeight = height;
            //.WrapText = true
            newWorksheet1.get_Range(a, Missing.Value).WrapText = iswarp;
        }
        public void SetMerg(string a1, string a2)
        {
            Microsoft.Office.Interop.Excel.Worksheet newWorksheet1 = Globals.ThisAddIn.Application.ActiveSheet;
            newWorksheet1.get_Range(a1, a2).Merge();
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox3.Text = (DaXie(textBox1.Text).Substring(DaXie(textBox1.Text).Length - 1, 1) == "元" ? DaXie(textBox1.Text) + "整" : DaXie(textBox1.Text));
        }
        public void Write(string path,string content)
        {
            FileStream fs = new FileStream(path, FileMode.Create);
            //获得字节数组
            byte[] data = System.Text.Encoding.Default.GetBytes(content);
            //开始写入
            fs.Write(data, 0, data.Length);
            //清空缓冲区、关闭流
            fs.Flush();
            fs.Close();
        }
       
        private void Form2_Load(object sender, EventArgs e)
        {
            textBox2.Text = DateTime.Now.ToString("D");
           textBox4.Text = File.ReadAllText("name.txt", Encoding.ASCII);

            load("khh.ini",comboBox1);
            load("fkdw.ini", comboBox2);
            load("zh.ini", comboBox3);
            load("skh.ini", comboBox4);
            load("skmc.ini", comboBox5);
            load("yt.ini", comboBox6);
            comboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox1.AutoCompleteSource = AutoCompleteSource.ListItems;
           // cmbTRADE_CO.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox2.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox2.AutoCompleteSource = AutoCompleteSource.ListItems;
            comboBox3.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox3.AutoCompleteSource = AutoCompleteSource.ListItems;
            comboBox4.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox4.AutoCompleteSource = AutoCompleteSource.ListItems;
            comboBox5.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox5.AutoCompleteSource = AutoCompleteSource.ListItems;
            comboBox6.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox6.AutoCompleteSource = AutoCompleteSource.ListItems;
        }
        public void load(string path,ComboBox c)
        {
            if (File.Exists( path))
            {
                StreamReader sr = null;
                try
                {
                    sr = new StreamReader( path, Encoding.UTF8);
                    //绑定内容到ComboBox
                    string item = sr.ReadLine();
                    while (item != null)
                    {if (item != "")
                        {
                            c.Items.Add(item);
                        }
                        item = sr.ReadLine();
                    }
                }
                catch (Exception ex)
                {
                   // MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (sr != null) sr.Close();
                }
            }
        }
        public void SetWidth(string col,Double width)
        {
            Microsoft.Office.Interop.Excel.Worksheet newWorksheet1 = Globals.ThisAddIn.Application.ActiveSheet;
            newWorksheet1.Columns[col +":"+ col, System.Type.Missing].ColumnWidth = width;
        }
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            textBox2.Text = DateTime.Now.ToString("D");
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            /*
             *   ADD(comboBox1,"khh.ini");

             ADD(comboBox2, "fkdw.ini");
             ADD(comboBox3, "zh.ini");

             ADD(comboBox4, "skh.ini");

             ADD(comboBox5, "skmc.ini");
             ADD(comboBox6, "yt.ini");

             Write("name.txt",textBox4.Text);
             */
            File.Delete("khh.ini");
            File.Delete("zh.ini");
            File.Delete("skh.ini");
            File.Delete("skmc.ini");
            File.Delete("yt.ini");
            MessageBox.Show("删除成功！");
        }
    }
}
