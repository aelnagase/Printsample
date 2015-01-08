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

using Excel = Microsoft.Office.Interop.Excel;       // 参照の追加 COM　"Microsoft Excel ??.? Object Library"
using System.Reflection;

namespace Printsample
{
    public partial class Form1 : Form
    {

        private string sExcelFile = @"C:\NagDevelop2\Printsample\test.xlsx";
        private Excel.Application oXls = null; // Excelオブジェクト
        private Excel.Workbook oWBook = null;  // workbookオブジェクト
        private Excel.Worksheet oSheet = null; // Worksheetオブジェクト

        private object oMissing = System.Reflection.Missing.Value;
        private object oTru = true;
        private object oFal = false;

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


        private void fnGetSheetList()
        {

            oXls = new Excel.Application();
            oXls.Visible = false; // 確認のためExcelのウィンドウを表示する?

            // Excelファイルをオープンする
            oWBook = (Excel.Workbook)(oXls.Workbooks.Open(
              sExcelFile,  // オープンするExcelファイル名
              Type.Missing, // （省略可能）UpdateLinks (0 / 1 / 2 / 3)
              oTru, // （省略可能）ReadOnly (True / False )
              Type.Missing, // （省略可能）Format
                // 1:タブ / 2:カンマ (,) / 3:スペース / 4:セミコロン (;)
                // 5:なし / 6:引数 Delimiterで指定された文字
              Type.Missing, // （省略可能）Password
              Type.Missing, // （省略可能）WriteResPassword
              Type.Missing, // （省略可能）IgnoreReadOnlyRecommended
              Type.Missing, // （省略可能）Origin
              Type.Missing, // （省略可能）Delimiter
              Type.Missing, // （省略可能）Editable
              Type.Missing, // （省略可能）Notify
              Type.Missing, // （省略可能）Converter
              Type.Missing, // （省略可能）AddToMru
              Type.Missing, // （省略可能）Local
              Type.Missing  // （省略可能）CorruptLoad
            ));

            // シート名の表示
            foreach (Excel.Worksheet sh in oWBook.Sheets)
            {
                Debug.Print(sh.Name);
            }

            oWBook.Close(Type.Missing, Type.Missing, Type.Missing);
            oXls.Quit();

            oWBook = null;
            oXls = null;

        }

        private void btnExcelSheetList_Click(object sender, EventArgs e)
        {
            fnGetSheetList();
        }

        private void btnExcleRead_Click(object sender, EventArgs e)
        {
            ReadTest();
        }


        private void ReadTest()
        {
            oXls = new Excel.Application();
            oXls.Visible = false; // 確認のためExcelのウィンドウを表示する

            // Excelファイルをオープンする
            oWBook = (Excel.Workbook)(oXls.Workbooks.Open(
              sExcelFile,  // オープンするExcelファイル名
              Type.Missing, // （省略可能）UpdateLinks (0 / 1 / 2 / 3)
              oTru, // （省略可能）ReadOnly (True / False )
              Type.Missing, // （省略可能）Format
                // 1:タブ / 2:カンマ (,) / 3:スペース / 4:セミコロン (;)
                // 5:なし / 6:引数 Delimiterで指定された文字
              Type.Missing, // （省略可能）Password
              Type.Missing, // （省略可能）WriteResPassword
              Type.Missing, // （省略可能）IgnoreReadOnlyRecommended
              Type.Missing, // （省略可能）Origin
              Type.Missing, // （省略可能）Delimiter
              Type.Missing, // （省略可能）Editable
              Type.Missing, // （省略可能）Notify
              Type.Missing, // （省略可能）Converter
              Type.Missing, // （省略可能）AddToMru
              Type.Missing, // （省略可能）Local
              Type.Missing  // （省略可能）CorruptLoad
            ));

            string sSheetName = "Sheet1";
            int nIndex = getSheetIndex(sSheetName, oWBook.Sheets);

            oSheet = (Excel.Worksheet)oWBook.Sheets[nIndex];

            string sNo = "", sName = "", sKana = "", sSex = "", sBirth = "", sBusyo = "";
            for (int nRow = 2; nRow < 1000; nRow++)
            {
                sNo = "";
                sNo = fnGetExcelItem(nRow, 1);
                if (sNo == "") break;
                sName  = fnGetExcelItem(nRow, 2);
                sKana  = fnGetExcelItem(nRow, 3);
                sSex   = fnGetExcelItem(nRow, 4);
                sBirth = fnGetExcelItem(nRow, 5);
                sBusyo = fnGetExcelItem(nRow, 6);
                Debug.Print(string.Format("{0} {1} {2} {3} {4} {5}", sNo , sName , sKana , sSex , sBirth , sBusyo));
            }

            oWBook.Close(Type.Missing, Type.Missing, Type.Missing);
            oXls.Quit();
        }


        private String fnGetExcelItem(int nRow, int nCol)
        {
            Excel.Range rng; // Rangeオブジェクト
            rng = (Excel.Range)oSheet.Cells[nRow, nCol];
            return (rng.Text.ToString());
        }


        private int fnGetValueExcelItem(int nRow, int nCol)
        {
            Excel.Range rng; // Rangeオブジェクト
            rng = (Excel.Range)oSheet.Cells[nRow, nCol];
            int nRet = 0;
            int.TryParse(rng.Value2.ToString(), out nRet);
            return (nRet);
        }



        // 指定されたワークシート名のインデックスを返すメソッド
        private int getSheetIndex(string sheetName, Excel.Sheets shs)
        {
            int i = 0;
            foreach (Excel.Worksheet sh in shs)
            {
                if (sheetName == sh.Name)
                {
                    return i + 1;
                }
                i += 1;
            }
            return 0;
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
