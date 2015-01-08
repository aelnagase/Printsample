using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

using System.Drawing.Printing;

namespace Printsample
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private string sTargetPrinterName = "";


        // 接続プリンターの一覧
        private void btnPrinterListUp_Click(object sender, EventArgs e)
        {
            foreach (string sPrinerName in PrinterSettings.InstalledPrinters)
            {
                Debug.Print(sPrinerName);
                if (sPrinerName.IndexOf("LBP9600C") >= 0)
                {
                    sTargetPrinterName = sPrinerName;
                }
            }
        }


        private void initializePrinterSetting1(string sPrinterName)
        {
            //プリンター名の設定
            if (sPrinterName != "")
            {
                this.printDocument1.PrinterSettings.PrinterName = sPrinterName;
                if (this.printDocument1.PrinterSettings.IsValid == false)
                {
                    //プリンター名異常
                    throw new Exception("プリンター名異常");
                }
            }

            // A4サイズを指定
            int nPaperIndex = 0;
            foreach (PaperSize ps in printDocument1.PrinterSettings.PaperSizes)
            {
                Debug.WriteLine(ps.PaperName);
                if (ps.PaperName.IndexOf("A4") > -1)
                {
                    break;
                }
                nPaperIndex++;
            }

            // 両面印刷可能のプリンター
            if (printDocument1.PrinterSettings.CanDuplex == true)
            {
                printDocument1.PrinterSettings.Duplex = Duplex.Simplex;　　//　片面印刷に変更
            }

            printDocument1.DefaultPageSettings.PaperSize = printDocument1.PrinterSettings.PaperSizes[nPaperIndex];
            // printDocument1.DefaultPageSettings.Landscape = true;    // 横向き
            printDocument1.DefaultPageSettings.Landscape = false;    // 縦向き
            //printDocument1.DefaultPageSettings.Color = true;        // カラー印刷
        }

        private void btnPrinTest_Click(object sender, EventArgs e)
        {
            initializePrinterSetting1(sTargetPrinterName);
   
            // プレビュー表示
            printPreviewDialog1.Document = printDocument1;
            printPreviewDialog1.ShowDialog();

            //　印刷
            // printDocument1.Print();

        }

        private printItem[] prtData1 = {
                new printItem(0.3F,39,28,100,37,"名前１",10,0,"name1"),
                new printItem(0F,39,37,100,48,"名前２",10,0,"name2"),
        };

        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {
            //　扱いやすい様に mm 単位スケールに変更
            e.Graphics.PageUnit = GraphicsUnit.Millimeter;

            float xMargn = -28;
            float yMargn = -5;

            for (int i = 0; i < prtData1.Length; i++)
            {
                Pen objPen;
                Font objFont;
                RectangleF oRect;
                StringFormat oFormat = new StringFormat();
                oFormat.Alignment = StringAlignment.Near;
                oFormat.LineAlignment = StringAlignment.Center;

                if (prtData1[i].PenWidth != 0)
                {
                    objPen = new Pen(Color.Black, (float)prtData1[i].PenWidth);
                    e.Graphics.DrawRectangle(objPen, prtData1[i].x + xMargn, prtData1[i].y + yMargn, prtData1[i].width, prtData1[i].height);
                }

                if (prtData1[i].sMessage != "")
                {
                    switch (prtData1[i].AlignMent)
                    {
                        case 0:
                            oFormat.Alignment = StringAlignment.Near;
                            break;
                        case 1:
                            oFormat.Alignment = StringAlignment.Center;
                            break;
                        case 2:
                            oFormat.Alignment = StringAlignment.Far;
                            break;
                    }

                    objFont = new Font("MS UI Gothic", prtData1[i].FontSize);
                    oRect = new RectangleF(prtData1[i].x + xMargn, prtData1[i].y + yMargn, prtData1[i].width, prtData1[i].height);
                    e.Graphics.DrawString(prtData1[i].sMessage, objFont, Brushes.Black, oRect, oFormat);
                }
            }

            // 多角形の描画
            { 
                int sx = 30;
                int sy = 60;
                int x1 = sx + 40;
                int x2 = sx + 80;
                int y1 = sy + 20;
                int y2 = sy + 60;
                int y3 = sy + 80;

                Point[] oPoints = {   
                    new Point(x1, sy),
                    new Point(sx, y1),
                    new Point(sx, y2),
                    new Point(x1, y3),
                    new Point(x2, y2),
                    new Point(x2, y1)
                };
                Pen objPen = new Pen(Color.Red, (float)0.3);
                e.Graphics.DrawPolygon(objPen, oPoints);

            }



            e.HasMorePages = false;
        }

       
    
    }

    public class printItem
    {
        public float PenWidth;
        public float x;
        public float y;
        public float width;
        public float height;
        public String sMessage;
        public int FontSize;
        public int AlignMent;
        public string sLabel;

        public printItem(float _PenWidth, float StartX, float StartY, float EndX, float EndY, String _sMessage, int _FontSize, int _AlignMent)
            : this(_PenWidth, StartX, StartY, EndX, EndY, _sMessage, _FontSize, _AlignMent, "")
        {
        }

        public printItem(float _PenWidth, float StartX, float StartY, float EndX, float EndY, String _sMessage, int _FontSize, int _AlignMent, string _sLabel)
        {
            PenWidth = _PenWidth;
            x = StartX;
            y = StartY;
            width = EndX - StartX;
            height = EndY - StartY;
            sMessage = _sMessage;
            FontSize = _FontSize;
            AlignMent = _AlignMent;
            sLabel = _sLabel;
        }

    }
}
