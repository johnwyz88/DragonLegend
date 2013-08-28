namespace ReceptionScreen
{
    partial class ReceptionScreen
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ReceptionScreen));
            this.pic_CompanyPicture = new System.Windows.Forms.PictureBox();
            this.lbl_CompanyName = new System.Windows.Forms.Label();
            this.btn_CreateTicket = new System.Windows.Forms.Button();
            this.lbl_NextTicket = new System.Windows.Forms.Label();
            this.lbl_ContactInfo = new System.Windows.Forms.Label();
            this.lbl_LabelNextTicket = new System.Windows.Forms.Label();
            this.lbl_LabelPrintTicket = new System.Windows.Forms.Label();
            this.lbl_DateTime = new System.Windows.Forms.Label();
            this.lbl_LabelPrintedNum = new System.Windows.Forms.Label();
            this.lbl_TopBreakLine = new System.Windows.Forms.Label();
            this.lbl_BottomBreakLine = new System.Windows.Forms.Label();
            this.WaitingList = new System.Windows.Forms.ListBox();
            this.lbl_PrintedNumber = new System.Windows.Forms.Label();
            this.tm_Update = new System.Windows.Forms.Timer(this.components);
            this.btn_Analysis = new System.Windows.Forms.Button();
            this.txb_NOF = new System.Windows.Forms.TextBox();
            this.pic_face = new System.Windows.Forms.PictureBox();
            this.lbl_waitingTickets = new System.Windows.Forms.Label();
            this.lbl_displayTotal = new System.Windows.Forms.Label();
            this.chk_printEnable = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.pic_CompanyPicture)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_face)).BeginInit();
            this.SuspendLayout();
            // 
            // pic_CompanyPicture
            // 
            this.pic_CompanyPicture.Image = ((System.Drawing.Image)(resources.GetObject("pic_CompanyPicture.Image")));
            this.pic_CompanyPicture.Location = new System.Drawing.Point(11, 11);
            this.pic_CompanyPicture.Margin = new System.Windows.Forms.Padding(2);
            this.pic_CompanyPicture.Name = "pic_CompanyPicture";
            this.pic_CompanyPicture.Size = new System.Drawing.Size(191, 125);
            this.pic_CompanyPicture.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pic_CompanyPicture.TabIndex = 2;
            this.pic_CompanyPicture.TabStop = false;
            // 
            // lbl_CompanyName
            // 
            this.lbl_CompanyName.AutoSize = true;
            this.lbl_CompanyName.Font = new System.Drawing.Font("Cooper Black", 49.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_CompanyName.Location = new System.Drawing.Point(235, 63);
            this.lbl_CompanyName.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_CompanyName.Name = "lbl_CompanyName";
            this.lbl_CompanyName.Size = new System.Drawing.Size(689, 77);
            this.lbl_CompanyName.TabIndex = 4;
            this.lbl_CompanyName.Text = "Dragon Legend Inc.";
            this.lbl_CompanyName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btn_CreateTicket
            // 
            this.btn_CreateTicket.Location = new System.Drawing.Point(550, 465);
            this.btn_CreateTicket.Margin = new System.Windows.Forms.Padding(2);
            this.btn_CreateTicket.Name = "btn_CreateTicket";
            this.btn_CreateTicket.Size = new System.Drawing.Size(78, 22);
            this.btn_CreateTicket.TabIndex = 8;
            this.btn_CreateTicket.Text = "Create ticket";
            this.btn_CreateTicket.UseVisualStyleBackColor = true;
            this.btn_CreateTicket.Click += new System.EventHandler(this.btn_CreateTicket_Click);
            // 
            // lbl_NextTicket
            // 
            this.lbl_NextTicket.AutoSize = true;
            this.lbl_NextTicket.Font = new System.Drawing.Font("MS Reference Sans Serif", 140.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_NextTicket.ForeColor = System.Drawing.Color.Red;
            this.lbl_NextTicket.Location = new System.Drawing.Point(439, 156);
            this.lbl_NextTicket.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_NextTicket.Name = "lbl_NextTicket";
            this.lbl_NextTicket.Size = new System.Drawing.Size(792, 229);
            this.lbl_NextTicket.TabIndex = 11;
            this.lbl_NextTicket.Text = "A000-0";
            // 
            // lbl_ContactInfo
            // 
            this.lbl_ContactInfo.AutoSize = true;
            this.lbl_ContactInfo.Location = new System.Drawing.Point(475, 621);
            this.lbl_ContactInfo.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_ContactInfo.Name = "lbl_ContactInfo";
            this.lbl_ContactInfo.Size = new System.Drawing.Size(192, 39);
            this.lbl_ContactInfo.TabIndex = 12;
            this.lbl_ContactInfo.Text = "              www.dragonlegend.ca\n                     905-940-1133\n25 Lanark Rd." +
    " Markham, ON, L3R 8E8";
            // 
            // lbl_LabelNextTicket
            // 
            this.lbl_LabelNextTicket.AutoSize = true;
            this.lbl_LabelNextTicket.Font = new System.Drawing.Font("MS Reference Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_LabelNextTicket.Location = new System.Drawing.Point(188, 300);
            this.lbl_LabelNextTicket.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_LabelNextTicket.Name = "lbl_LabelNextTicket";
            this.lbl_LabelNextTicket.Size = new System.Drawing.Size(129, 19);
            this.lbl_LabelNextTicket.TabIndex = 13;
            this.lbl_LabelNextTicket.Text = "Next Ticket # ";
            // 
            // lbl_LabelPrintTicket
            // 
            this.lbl_LabelPrintTicket.AutoSize = true;
            this.lbl_LabelPrintTicket.Font = new System.Drawing.Font("MS Reference Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_LabelPrintTicket.Location = new System.Drawing.Point(188, 465);
            this.lbl_LabelPrintTicket.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_LabelPrintTicket.Name = "lbl_LabelPrintTicket";
            this.lbl_LabelPrintTicket.Size = new System.Drawing.Size(211, 19);
            this.lbl_LabelPrintTicket.TabIndex = 14;
            this.lbl_LabelPrintTicket.Text = "Number of People    -->";
            // 
            // lbl_DateTime
            // 
            this.lbl_DateTime.AutoSize = true;
            this.lbl_DateTime.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_DateTime.Location = new System.Drawing.Point(1250, 645);
            this.lbl_DateTime.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_DateTime.Name = "lbl_DateTime";
            this.lbl_DateTime.Size = new System.Drawing.Size(67, 15);
            this.lbl_DateTime.TabIndex = 15;
            this.lbl_DateTime.Text = "Date & Time";
            // 
            // lbl_LabelPrintedNum
            // 
            this.lbl_LabelPrintedNum.AutoSize = true;
            this.lbl_LabelPrintedNum.Font = new System.Drawing.Font("MS Reference Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_LabelPrintedNum.Location = new System.Drawing.Point(189, 519);
            this.lbl_LabelPrintedNum.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_LabelPrintedNum.Name = "lbl_LabelPrintedNum";
            this.lbl_LabelPrintedNum.Size = new System.Drawing.Size(142, 19);
            this.lbl_LabelPrintedNum.TabIndex = 16;
            this.lbl_LabelPrintedNum.Text = "Printed Ticket #";
            // 
            // lbl_TopBreakLine
            // 
            this.lbl_TopBreakLine.AutoSize = true;
            this.lbl_TopBreakLine.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.lbl_TopBreakLine.Location = new System.Drawing.Point(-37, 169);
            this.lbl_TopBreakLine.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_TopBreakLine.Name = "lbl_TopBreakLine";
            this.lbl_TopBreakLine.Size = new System.Drawing.Size(2050, 13);
            this.lbl_TopBreakLine.TabIndex = 18;
            this.lbl_TopBreakLine.Text = resources.GetString("lbl_TopBreakLine.Text");
            // 
            // lbl_BottomBreakLine
            // 
            this.lbl_BottomBreakLine.AutoSize = true;
            this.lbl_BottomBreakLine.ForeColor = System.Drawing.Color.CornflowerBlue;
            this.lbl_BottomBreakLine.Location = new System.Drawing.Point(-346, 569);
            this.lbl_BottomBreakLine.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_BottomBreakLine.Name = "lbl_BottomBreakLine";
            this.lbl_BottomBreakLine.Size = new System.Drawing.Size(2050, 13);
            this.lbl_BottomBreakLine.TabIndex = 21;
            this.lbl_BottomBreakLine.Text = resources.GetString("lbl_BottomBreakLine.Text");
            // 
            // WaitingList
            // 
            this.WaitingList.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.WaitingList.FormattingEnabled = true;
            this.WaitingList.Location = new System.Drawing.Point(1292, 185);
            this.WaitingList.Name = "WaitingList";
            this.WaitingList.Size = new System.Drawing.Size(53, 420);
            this.WaitingList.TabIndex = 22;
            this.WaitingList.SelectedIndexChanged += new System.EventHandler(this.WaitingList_DoubleClick);
            // 
            // lbl_PrintedNumber
            // 
            this.lbl_PrintedNumber.AutoSize = true;
            this.lbl_PrintedNumber.BackColor = System.Drawing.Color.Lime;
            this.lbl_PrintedNumber.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_PrintedNumber.Location = new System.Drawing.Point(483, 519);
            this.lbl_PrintedNumber.Name = "lbl_PrintedNumber";
            this.lbl_PrintedNumber.Size = new System.Drawing.Size(81, 25);
            this.lbl_PrintedNumber.TabIndex = 23;
            this.lbl_PrintedNumber.Text = "A000-0";
            // 
            // tm_Update
            // 
            this.tm_Update.Enabled = true;
            this.tm_Update.Interval = 1000;
            this.tm_Update.Tick += new System.EventHandler(this.tm_Update_Tick);
            // 
            // btn_Analysis
            // 
            this.btn_Analysis.Location = new System.Drawing.Point(95, 637);
            this.btn_Analysis.Name = "btn_Analysis";
            this.btn_Analysis.Size = new System.Drawing.Size(60, 23);
            this.btn_Analysis.TabIndex = 25;
            this.btn_Analysis.Text = "Analysis";
            this.btn_Analysis.UseVisualStyleBackColor = true;
            this.btn_Analysis.Click += new System.EventHandler(this.btn_Analysis_Click);
            // 
            // txb_NOF
            // 
            this.txb_NOF.Location = new System.Drawing.Point(488, 466);
            this.txb_NOF.Name = "txb_NOF";
            this.txb_NOF.Size = new System.Drawing.Size(25, 20);
            this.txb_NOF.TabIndex = 26;
            // 
            // pic_face
            // 
            this.pic_face.Image = global::ReceptionScreen.Properties.Resources.Opera_face_Gif3;
            this.pic_face.Location = new System.Drawing.Point(1238, 11);
            this.pic_face.Name = "pic_face";
            this.pic_face.Size = new System.Drawing.Size(99, 125);
            this.pic_face.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pic_face.TabIndex = 27;
            this.pic_face.TabStop = false;
            // 
            // lbl_waitingTickets
            // 
            this.lbl_waitingTickets.AutoSize = true;
            this.lbl_waitingTickets.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_waitingTickets.ForeColor = System.Drawing.Color.Crimson;
            this.lbl_waitingTickets.Location = new System.Drawing.Point(1151, 521);
            this.lbl_waitingTickets.Name = "lbl_waitingTickets";
            this.lbl_waitingTickets.Size = new System.Drawing.Size(194, 48);
            this.lbl_waitingTickets.TabIndex = 29;
            this.lbl_waitingTickets.Text = "| A : 0 | B : 0 | C : 0 |\r\n            -> Total : 0";
            // 
            // lbl_displayTotal
            // 
            this.lbl_displayTotal.AutoSize = true;
            this.lbl_displayTotal.Font = new System.Drawing.Font("Monotype Corsiva", 20.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_displayTotal.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.lbl_displayTotal.Location = new System.Drawing.Point(1147, 579);
            this.lbl_displayTotal.Name = "lbl_displayTotal";
            this.lbl_displayTotal.Size = new System.Drawing.Size(208, 66);
            this.lbl_displayTotal.TabIndex = 30;
            this.lbl_displayTotal.Text = "Total tickets: ???\r\nTotal customers: ???";
            // 
            // chk_printEnable
            // 
            this.chk_printEnable.AutoSize = true;
            this.chk_printEnable.Location = new System.Drawing.Point(657, 467);
            this.chk_printEnable.Name = "chk_printEnable";
            this.chk_printEnable.Size = new System.Drawing.Size(86, 17);
            this.chk_printEnable.TabIndex = 31;
            this.chk_printEnable.Text = "PrintEnabled";
            this.chk_printEnable.UseVisualStyleBackColor = true;
            // 
            // ReceptionScreen
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Lavender;
            this.ClientSize = new System.Drawing.Size(1350, 729);
            this.Controls.Add(this.chk_printEnable);
            this.Controls.Add(this.lbl_displayTotal);
            this.Controls.Add(this.lbl_waitingTickets);
            this.Controls.Add(this.pic_face);
            this.Controls.Add(this.txb_NOF);
            this.Controls.Add(this.btn_Analysis);
            this.Controls.Add(this.lbl_PrintedNumber);
            this.Controls.Add(this.WaitingList);
            this.Controls.Add(this.lbl_BottomBreakLine);
            this.Controls.Add(this.lbl_TopBreakLine);
            this.Controls.Add(this.lbl_LabelPrintedNum);
            this.Controls.Add(this.lbl_DateTime);
            this.Controls.Add(this.lbl_LabelPrintTicket);
            this.Controls.Add(this.lbl_LabelNextTicket);
            this.Controls.Add(this.lbl_ContactInfo);
            this.Controls.Add(this.lbl_NextTicket);
            this.Controls.Add(this.btn_CreateTicket);
            this.Controls.Add(this.lbl_CompanyName);
            this.Controls.Add(this.pic_CompanyPicture);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "ReceptionScreen";
            this.Text = "ReceptionScreen";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.ReceptionScreen_Load);
            this.MouseMove += new System.Windows.Forms.MouseEventHandler(this.ReceptionScreen_MouseMove);
            ((System.ComponentModel.ISupportInitialize)(this.pic_CompanyPicture)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_face)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pic_CompanyPicture;
        private System.Windows.Forms.Label lbl_CompanyName;
        private System.Windows.Forms.Button btn_CreateTicket;
        private System.Windows.Forms.Label lbl_NextTicket;
        private System.Windows.Forms.Label lbl_ContactInfo;
        private System.Windows.Forms.Label lbl_LabelNextTicket;
        private System.Windows.Forms.Label lbl_LabelPrintTicket;
        private System.Windows.Forms.Label lbl_DateTime;
        private System.Windows.Forms.Label lbl_LabelPrintedNum;
        private System.Windows.Forms.Label lbl_TopBreakLine;
        private System.Windows.Forms.Label lbl_BottomBreakLine;
        private System.Windows.Forms.ListBox WaitingList;
        private System.Windows.Forms.Label lbl_PrintedNumber;
        private System.Windows.Forms.Timer tm_Update;
        private System.Windows.Forms.Button btn_Analysis;
        private System.Windows.Forms.TextBox txb_NOF;
        private System.Windows.Forms.PictureBox pic_face;
        private System.Windows.Forms.Label lbl_waitingTickets;
        private System.Windows.Forms.Label lbl_displayTotal;
        private System.Windows.Forms.CheckBox chk_printEnable;
    }
}

