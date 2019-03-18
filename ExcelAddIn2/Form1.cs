using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ExcelAddIn2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
        }
         public void method(object str)
        {
            button1.Text = "正在生成...";
            Microsoft.Office.Interop.Excel._Workbook wbook = Globals.ThisAddIn.Application.ActiveWorkbook;//获取激活的工作簿
                                                                                                          //  group1.Label = "当前共有 "+wbook.Sheets.Count.ToString()+" 张表\r\n自动引导至第一张表，表名为："+wbook.Sheets[1].Name;//获取第一个工作表;

            //  Microsoft.Office.Interop.Excel.Worksheet newWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            // newWorksheet.Name = "Sheet1";
            Worksheet worksheet = wbook.Worksheets[1];//获取名为sheet1的工作表
          
            Microsoft.Office.Interop.Excel.Worksheet newWorkbook = Globals.ThisAddIn.Application.Worksheets.Add(System.Type.Missing, worksheet);
            //  Microsoft.Office.Interop.Excel.Worksheet newWorkbook = Globals.ThisAddIn.Application.Worksheets.Add(System.Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet newWorksheet1 = Globals.ThisAddIn.Application.ActiveSheet;
            newWorksheet1.Columns["A:A", System.Type.Missing].ColumnWidth = 6.5;
            newWorksheet1.Columns["B:B", System.Type.Missing].ColumnWidth = 8;                //  Thread thread = new Thread(new ParameterizedThreadStart(method));//创建线程
            newWorksheet1.Columns["C:C", System.Type.Missing].ColumnWidth = 10.38;
            newWorksheet1.Columns["D:D", System.Type.Missing].ColumnWidth = 13;
            newWorksheet1.Columns["E:E", System.Type.Missing].ColumnWidth = 28.38;
            newWorksheet1.Columns["F:F", System.Type.Missing].ColumnWidth = 6.5;
            newWorksheet1.Columns["G:G", System.Type.Missing].ColumnWidth = 14.63;
            newWorksheet1.PageSetup.TopMargin = 1.9;
            newWorksheet1.PageSetup.BottomMargin = 1.9;
            newWorksheet1.PageSetup.LeftMargin = 0.7;
            newWorksheet1.PageSetup.RightMargin = 1.3;
            
            for (int I = 4; I <=worksheet.UsedRange.Rows.Count; I++)
            {
            progressBar1.Value= (I*100)/(worksheet.UsedRange.Rows.Count);
                newWorksheet1.get_Range("A" + (3 + (I - 4) * 9), Missing.Value).Value2 = "余额对账单(对公)";//设置某个单元格的值
                newWorksheet1.get_Range("A" + (3 + (I - 4) * 9), Missing.Value).Font.Size = 10;
                newWorksheet1.get_Range("A" + (3 + (I - 4) * 9), Missing.Value).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                //newWorksheet1.get_Range("G"+(5 + (I - 4) * 9), Missing.Value).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                //   newWorksheet1.get_Range("A3", Missing.Value).VerticalAlignment =ve;
                //青阳县大华丝绸有限责任公司
                newWorksheet1.get_Range("A" + (4 + (I - 4) * 9), Missing.Value).Value2 = worksheet.Cells[I, 2];
                newWorksheet1.get_Range("A" + (4 + (I - 4) * 9), Missing.Value).Font.Size = 10;
                newWorksheet1.get_Range("D" + (4 + (I - 4) * 9), Missing.Value).Value2 = ":";
                newWorksheet1.get_Range("D" + (4 + (I - 4) * 9), Missing.Value).Font.Size = 10;
                newWorksheet1.get_Range("A" + (3 + (I - 4) * 9), "G" + (3 + (I - 4) * 9)).Merge();
                newWorksheet1.get_Range("A" + (4+ (I - 4) * 9), "C" + (4 + (I - 4) * 9)).Merge();
                newWorksheet1.get_Range("C" + (5 + (I - 4) * 9), "D" + (5 + (I - 4) * 9)).Merge();
                newWorksheet1.get_Range("D" + (6 + (I - 4) * 9), "E" + (6 + (I - 4) * 9)).Merge();
                newWorksheet1.get_Range("B" + (7 + (I - 4) * 9), "G" + (7 + (I - 4) * 9)).Merge();
                newWorksheet1.get_Range("C" + (5 + (I - 4) * 9), Missing.Value).Value2 = "贵单位在我处存款账号:";
                newWorksheet1.get_Range("C" + (5 + (I - 4) * 9), Missing.Value).Font.Size = 10;
                newWorksheet1.get_Range("E" + (5 + (I - 4) * 9), Missing.Value).Value2 = worksheet.Cells[I, 1];
                newWorksheet1.get_Range("E" + (5 + (I - 4) * 9), Missing.Value).Font.Size = 10;
                newWorksheet1.get_Range("F" + (5 + (I - 4) * 9), Missing.Value).Value2 = "，截止";
                newWorksheet1.get_Range("F" + (5 + (I - 4) * 9), Missing.Value).Font.Size = 10;
                newWorksheet1.get_Range("G" + (5 + (I - 4) * 9), Missing.Value).Value2 = "'"+textBox1.Text ;
                newWorksheet1.get_Range("G" + (5 + (I - 4) * 9), Missing.Value).Font.Size = 10;
                newWorksheet1.get_Range("B" + (6 + (I - 4) * 9), Missing.Value).Value2 = "余额为：";
                newWorksheet1.get_Range("B" + (6 + (I - 4) * 9), Missing.Value).Font.Size = 10;
                newWorksheet1.get_Range("C" + (6 + (I - 4) * 9), Missing.Value).Value2 = worksheet.Cells[I, 3];
                newWorksheet1.get_Range("C" + (6 + (I - 4) * 9), Missing.Value).Font.Size = 10;
                //thread.Start(3);
                //，随附对账明细。请即核对，如有不符，速来核查。//核对后请填写回复联，并加盖续留印签，于10日内返回我社（如有未达账务，请逐笔填列未达账务调节表）。					
                newWorksheet1.get_Range("D" + (6 + (I - 4) * 9), Missing.Value).Value2 = "，随附对账明细。请即核对，如有不符，速来核查。";
                newWorksheet1.get_Range("D" + (6 + (I - 4) * 9), Missing.Value).Font.Size = 10;
                newWorksheet1.get_Range("B" + (7 + (I - 4) * 9), Missing.Value).Value2 = "核对后请填写回复联，并加盖预留印签，于10日内返回我社（如有未达账务，请逐笔填列未达账务调节表）。	";
                newWorksheet1.get_Range("B" + (7 + (I - 4) * 9), Missing.Value).Font.Size = 10;
                newWorksheet1.get_Range("B" + (8 + (I - 4) * 9), Missing.Value).Value2 = "此敬！";
                newWorksheet1.get_Range("B" + (8 + (I - 4) * 9), Missing.Value).Font.Size = 10;
                newWorksheet1.get_Range("E" + (9 + (I - 4) * 9), Missing.Value).Value2 = "'" + textBox2.Text;
                newWorksheet1.get_Range("E" + (9 + (I - 4) * 9), Missing.Value).Font.Size = 10;
                newWorksheet1.HPageBreaks.Add(newWorksheet1.get_Range("A" + (10+ (I - 4) * 9), Missing.Value));
            }
            //   group1.Label = "生成结束，请检查完整性。感谢您的使用！";
            MessageBox.Show("生成结束，请检查完整性。感谢您的使用！", "提示");
            this.Close();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "")
            {
                Random ran = new Random();
                //保存当前窗体位置
                System.Drawing.Point oldPoint = this.Location;
                for (int i = 0; i < 15; i++)
                {
                    //随机生成新的位置
                    System.Drawing.Point newPoint = new System.Drawing.Point(oldPoint.X + ran.Next(-10, 10), oldPoint.Y + ran.Next(-10, 10));
                    //将位置设置给窗体
                    this.Location = newPoint;
                    Thread.Sleep(10);
                    this.Location = oldPoint;
                    //休息50毫秒
                   // Thread.Sleep(50);
                }
              
                return;
            }
            textBox1.Enabled = false;
            textBox2.Enabled = false;
            button1.Enabled = false;
            Thread thread1 = new Thread(new ParameterizedThreadStart(method));
            thread1.Start();
        }
    }
}
