using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace 음료관리프로그램
{
    public partial class Form1 : Form
    {
        Microsoft.Office.Interop.Excel.Application excelapp = new Microsoft.Office.Interop.Excel.Application();
        Workbook wb = null;
        Worksheet ws = null;
        static int idx;
        static int totalprice = 0;
        Drink[] drinks =
           {
            new Drink("시트라", 2000),
            new Drink("울트라", 2000),
            new Drink("파라다이스", 2000),
            new Drink("포카리", 2000),
            new Drink("하늘보리", 1500),
            new Drink("삼다수", 1000),
            new Drink("핫식스", 1500),
            new Drink("씨그램", 1500),
            new Drink("미에로화이바", 1800),
            new Drink("트레비", 1500),
            new Drink("아티제", 1000),
            new Drink("제로콜라", 2000),
        };
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            wb = excelapp.Workbooks.Open(Filename: @"C:\Users\sna14\OneDrive\바탕 화면\ReStudy\C#\음료관리프로그램\음료관리프로그램\bin\2022년 음료판매일지.xlsx");

            ws = wb.Worksheets.Item["6월"];
            idx=ws.UsedRange.Rows.Count+1;

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            wb.Close();
            excelapp.Quit();

            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(excelapp);
        }
        public void Func(int idx)
        {
            listView1.Items.Add(new ListViewItem(new string[] { drinks[idx].Name, drinks[idx].Price.ToString() }));
            totalprice += drinks[idx].Price;
            textBox1.Text = totalprice.ToString();
        }
        public void ExelWrite(string how)
        {
            try
            {
                
               
                for (int count = 0; count < listView1.Items.Count; count++)
                {

                    Range rg = ws.Cells[idx, 1];
                    rg.Value = drinks[0].Dt.ToString("M월d일");

                    
                    rg = ws.Cells[idx, 2];
                    rg.Value = listView1.Items[count].SubItems[0].Text;
                    

                    rg = ws.Cells[idx, 3];
                    rg.Value = "1";
                    

                    rg = ws.Cells[idx, 4];
                    rg.Value = how;
              

                    rg = ws.Cells[idx, 5];
                    rg.Value = listView1.Items[count].SubItems[1].Text;

                    idx++;
                 
                }

                wb.Save();

                totalprice = 0;
                listView1.Items.Clear();
                textBox1.Text = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        
        #region 결제버튼
        private void 현금_Click(object sender, EventArgs e)
        {
            ExelWrite("현금");
        }

        private void 카드_Click(object sender, EventArgs e)
        {
            ExelWrite("카드");
        }

        private void 계좌_Click(object sender, EventArgs e)
        {
            ExelWrite("현금");
        }
        #endregion

        #region 음료버튼
        private void 시트라_Click(object sender, EventArgs e)
        {
            Func(0);
        }

        private void 울트라_Click(object sender, EventArgs e)
        {
            Func(1);
        }

        private void 파라다이스_Click(object sender, EventArgs e)
        {
            Func(2);
        }

        private void 포카리_Click(object sender, EventArgs e)
        {
            Func(3);
        }

        private void 하늘보리_Click(object sender, EventArgs e)
        {
            Func(4);
        }

        private void 삼다수_Click(object sender, EventArgs e)
        {
            Func(5);
        }

        private void 핫식스_Click(object sender, EventArgs e)
        {
            Func(6);
        }

        private void 씨그램_Click(object sender, EventArgs e)
        {
            Func(7);
        }

        private void 미에로화이바_Click(object sender, EventArgs e)
        {
            Func(8);
        }

        private void 트레비_Click(object sender, EventArgs e)
        {
            Func(9);
        }

        private void 아티제_Click(object sender, EventArgs e)
        {
            Func(10);
        }

        private void 제로콜라_Click(object sender, EventArgs e)
        {
            Func(11);
        }

        #endregion

        private void 취소_Click(object sender, EventArgs e)
        {
            totalprice = 0;
            listView1.Items.Clear();
            textBox1.Text = "";
        }

    }
}
