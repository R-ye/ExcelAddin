using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Reflection;
using System.Threading;
using System.Drawing;

namespace ExcelAddIn2
{
    public partial class Ribbon1
    {
        cutter cutter = null;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            Control.CheckForIllegalCrossThreadCalls = false;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Excel._Workbook wbook = Globals.ThisAddIn.Application.ActiveWorkbook;//获取激活的工作簿
                                                                                                          //  group1.Label = "当前共有 "+wbook.Sheets.Count.ToString()+" 张表\r\n自动引导至第一张表，表名为："+wbook.Sheets[1].Name;//获取第一个工作表;

            //  Microsoft.Office.Interop.Excel.Worksheet newWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            // newWorksheet.Name = "Sheet1";
            Worksheet worksheet = wbook.Worksheets[1];//获取名为sheet1的工作表
            string tag = worksheet.get_Range("A1", Missing.Value).Value2+"###";
            if (tag.Length <=0 || tag.IndexOf("余额表")==-1)
            {
                MessageBox.Show("未识别到有效数据","错误");
                return;
            }
            Form1 f = new Form1();
            f.ShowDialog();


        }
       
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
           
        }

        private void button2_Click_1(object sender, RibbonControlEventArgs e)
        {
           
        }

        private void button2_Click_2(object sender, RibbonControlEventArgs e)
        {
                   }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("不能重复进行生成操作，多线程状态下，会导致Excel发生APPCRASH错误。","注意事项");
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("本插件使用C#开发，源代码版权归开发者所有。", "关于插件");
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button2_Click_3(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Excel._Workbook wbook = Globals.ThisAddIn.Application.ActiveWorkbook;//获取激活的工作簿
                                                                                                          //  group1.Label = "当前共有 "+wbook.Sheets.Count.ToString()+" 张表\r\n自动引导至第一张表，表名为："+wbook.Sheets[1].Name;//获取第一个工作表;

            //  Microsoft.Office.Interop.Excel.Worksheet newWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            // newWorksheet.Name = "Sheet1";
            Worksheet worksheet = wbook.Worksheets[1];//获取名为sheet1的工作表
         //   Worksheet worksheet = wbook.Worksheets[1];//获取名为sheet1的工作表
            string tag = worksheet.get_Range("A1", Missing.Value).Value2+"###";
            if (tag.Length <= 0 || tag.IndexOf("日计表") == -1)
            {
                MessageBox.Show("未识别到有效数据", "错误");
                return;
            }

            int num = worksheet.Range["A1:H"+ worksheet.UsedRange.Rows.Count].Find("100103000000").Row; //第一次出现空单元格的行数
            string str1 = worksheet.get_Range("E" + num, Missing.Value).Value2;
            string str2 = worksheet.get_Range("F" + num, Missing.Value).Value2;
            string str3 = worksheet.get_Range("G" + num, Missing.Value).Value2;
            int num1 = worksheet.Range["A1:H" + worksheet.UsedRange.Rows.Count].Find("100102000000").Row; //第一次出现空单元格的行数
            string str11 = worksheet.get_Range("E" + num1, Missing.Value).Value2;
            string str21 = worksheet.get_Range("F" + num1, Missing.Value).Value2;
            string str31 = worksheet.get_Range("G" + num1, Missing.Value).Value2;
            double sum = double.Parse(str1) + double.Parse(str11);
            double sum1 = double.Parse(str2) + double.Parse(str21);
            double sum2 = double.Parse(str3) + double.Parse(str31);
          //  MessageBox.Show(sum.ToString());
           MessageBox.Show("     现金收付\r\n----------------------------\r\n     收入："+(sum).ToString()+"\r\n     付出：" + (sum1).ToString() + "\r\n     余额："+ (sum2).ToString());
        }

        private void button3_Click_1(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("1.请从网页系统中下载<对公余额表>的Excel文件\r\n2.使用不高于2016版本的Excel打开已经下载的对公余额表文件\r\n3.输入对账日期，即落款日期；输入截止日期，对账余额的计算终止日期\r\n4.点击“一键生成对账单”\r\n5.待提示生成结束，即可点击打印。", "使用说明");

        }

        private void button4_Click_1(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("1.请从网页系统中下载<日计表>的Excel文件\r\n2.使用不高于2016版本的Excel打开已经下载的日计表文件\r\n3.插件会自动计算出日计表生成日期的收付及余额数据。", "使用说明");

        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("本插件使用C#开发，插件使用了多线程，使用时请避免多开使用，防止发生错误。更多插件内容敬请期待。\r\n开发者：Oneday", "关于插件");
        }

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            Form2 f = new Form2();
            f.ShowDialog();
        }

        private void button6_Click_1(object sender, RibbonControlEventArgs e)
        {
            Form3 f = new Form3();
            f.ShowDialog();
        }

        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("本工具可以匹配任何工资表中的有效数据，使用时可先手动剔除无关数据，也可完全依靠程序剔除，剔除完成后，请再人工与原始数据进行对比，尤其是组数和金额总数，当发现不一致时，请根据程序提示进行处理，正常情况下可以提取全部有效数据。本程序对无效数据的定义是：缺少账号、姓名、金额中任意一项或多项均为无效数据，在存在序号时，会对缺少金额的数据进行标红，以便于联系工资单提供方。","代发工资提取工具使用说明");
        }

        private void button9_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("calc.exe");
        }

        private void button10_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("notepad.exe");
        }

        private void button11_Click(object sender, RibbonControlEventArgs e)
        {
            Form4 f = new Form4();
            f.Show();
        }

        private void button12_Click(object sender, RibbonControlEventArgs e)
        {
            Bitmap CatchBmp = new Bitmap(Screen.AllScreens[0].Bounds.Width, Screen.AllScreens[0].Bounds.Height);

            // 创建一个画板，让我们可以在画板上画图
            // 这个画板也就是和屏幕大小一样大的图片
            // 我们可以通过Graphics这个类在这个空白图片上画图
            Graphics g = Graphics.FromImage(CatchBmp);

            // 把屏幕图片拷贝到我们创建的空白图片 CatchBmp中
            g.CopyFromScreen(new System.Drawing.Point(0, 0), new System.Drawing.Point(0, 0), new Size(Screen.AllScreens[0].Bounds.Width, Screen.AllScreens[0].Bounds.Height));

            // 创建截图窗体
            cutter = new cutter();

            // 指示窗体的背景图片为屏幕图片
            cutter.BackgroundImage = CatchBmp;
            // 显示窗体
            //cutter.Show();
            // 如果Cutter窗体结束，则从剪切板获得截取的图片，并显示在聊天窗体的发送框中
            if (cutter.ShowDialog() == DialogResult.OK)
            {
                IDataObject iData = Clipboard.GetDataObject();
                DataFormats.Format format = DataFormats.GetFormat(DataFormats.Bitmap);
                if (iData.GetDataPresent(DataFormats.Bitmap))
                {
                  //  richTextBox1.Paste(format);

                    // 清楚剪贴板的图片
                    //Clipboard.Clear();
                }
            }
        }
    }
}
