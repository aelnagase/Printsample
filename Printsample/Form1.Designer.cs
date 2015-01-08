namespace Printsample
{
    partial class Form1
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースが破棄される場合 true、破棄されない場合は false です。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.btnPrinterListUp = new System.Windows.Forms.Button();
            this.printDocument1 = new System.Drawing.Printing.PrintDocument();
            this.printPreviewDialog1 = new System.Windows.Forms.PrintPreviewDialog();
            this.btnPrinTest = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnPrinterListUp
            // 
            this.btnPrinterListUp.Location = new System.Drawing.Point(63, 40);
            this.btnPrinterListUp.Name = "btnPrinterListUp";
            this.btnPrinterListUp.Size = new System.Drawing.Size(150, 23);
            this.btnPrinterListUp.TabIndex = 0;
            this.btnPrinterListUp.Text = "接続プリンター一覧";
            this.btnPrinterListUp.UseVisualStyleBackColor = true;
            this.btnPrinterListUp.Click += new System.EventHandler(this.btnPrinterListUp_Click);
            // 
            // printDocument1
            // 
            this.printDocument1.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.printDocument1_PrintPage);
            // 
            // printPreviewDialog1
            // 
            this.printPreviewDialog1.AutoScrollMargin = new System.Drawing.Size(0, 0);
            this.printPreviewDialog1.AutoScrollMinSize = new System.Drawing.Size(0, 0);
            this.printPreviewDialog1.ClientSize = new System.Drawing.Size(400, 300);
            this.printPreviewDialog1.Enabled = true;
            this.printPreviewDialog1.Icon = ((System.Drawing.Icon)(resources.GetObject("printPreviewDialog1.Icon")));
            this.printPreviewDialog1.Name = "printPreviewDialog1";
            this.printPreviewDialog1.Visible = false;
            // 
            // btnPrinTest
            // 
            this.btnPrinTest.Location = new System.Drawing.Point(67, 120);
            this.btnPrinTest.Name = "btnPrinTest";
            this.btnPrinTest.Size = new System.Drawing.Size(150, 23);
            this.btnPrinTest.TabIndex = 1;
            this.btnPrinTest.Text = "印刷テスト";
            this.btnPrinTest.UseVisualStyleBackColor = true;
            this.btnPrinTest.Click += new System.EventHandler(this.btnPrinTest_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 262);
            this.Controls.Add(this.btnPrinTest);
            this.Controls.Add(this.btnPrinterListUp);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnPrinterListUp;
        private System.Drawing.Printing.PrintDocument printDocument1;
        private System.Windows.Forms.PrintPreviewDialog printPreviewDialog1;
        private System.Windows.Forms.Button btnPrinTest;

    }
}

