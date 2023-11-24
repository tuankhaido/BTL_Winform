namespace WindowsFormsApp3
{
    partial class Form1
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
            this.label1 = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.lớpHọcToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.danhSáchLớpHọcToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.báoCáoDoanhThuTheoLớpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.báoCáoDanhSáchLớpTheoPhòngToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.họcViênToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.danhSáchHọcViênToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.báoCáoĐiểmTheoLớpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.báoCáoCácHọcViênĐiểmCaoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.giáoViênToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.danhSáchGiáoViênToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(44, 45);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(708, 32);
            this.label1.TabIndex = 0;
            this.label1.Text = "Quản lý học viên lớp tin học và ngoại ngữ ngắn hạn";
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.lớpHọcToolStripMenuItem,
            this.họcViênToolStripMenuItem,
            this.giáoViênToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(800, 28);
            this.menuStrip1.TabIndex = 2;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // lớpHọcToolStripMenuItem
            // 
            this.lớpHọcToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.danhSáchLớpHọcToolStripMenuItem,
            this.báoCáoDoanhThuTheoLớpToolStripMenuItem,
            this.báoCáoDanhSáchLớpTheoPhòngToolStripMenuItem});
            this.lớpHọcToolStripMenuItem.Name = "lớpHọcToolStripMenuItem";
            this.lớpHọcToolStripMenuItem.Size = new System.Drawing.Size(76, 24);
            this.lớpHọcToolStripMenuItem.Text = "Lớp học";
            // 
            // danhSáchLớpHọcToolStripMenuItem
            // 
            this.danhSáchLớpHọcToolStripMenuItem.Name = "danhSáchLớpHọcToolStripMenuItem";
            this.danhSáchLớpHọcToolStripMenuItem.Size = new System.Drawing.Size(323, 26);
            this.danhSáchLớpHọcToolStripMenuItem.Text = "Danh sách lớp học";
            this.danhSáchLớpHọcToolStripMenuItem.Click += new System.EventHandler(this.danhSáchLớpHọcToolStripMenuItem_Click);
            // 
            // báoCáoDoanhThuTheoLớpToolStripMenuItem
            // 
            this.báoCáoDoanhThuTheoLớpToolStripMenuItem.Name = "báoCáoDoanhThuTheoLớpToolStripMenuItem";
            this.báoCáoDoanhThuTheoLớpToolStripMenuItem.Size = new System.Drawing.Size(323, 26);
            this.báoCáoDoanhThuTheoLớpToolStripMenuItem.Text = "Báo cáo doanh thu theo lớp";
            this.báoCáoDoanhThuTheoLớpToolStripMenuItem.Click += new System.EventHandler(this.báoCáoDoanhThuTheoLớpToolStripMenuItem_Click);
            // 
            // báoCáoDanhSáchLớpTheoPhòngToolStripMenuItem
            // 
            this.báoCáoDanhSáchLớpTheoPhòngToolStripMenuItem.Name = "báoCáoDanhSáchLớpTheoPhòngToolStripMenuItem";
            this.báoCáoDanhSáchLớpTheoPhòngToolStripMenuItem.Size = new System.Drawing.Size(323, 26);
            this.báoCáoDanhSáchLớpTheoPhòngToolStripMenuItem.Text = "Báo cáo danh sách lớp theo phòng";
            this.báoCáoDanhSáchLớpTheoPhòngToolStripMenuItem.Click += new System.EventHandler(this.báoCáoDanhSáchLớpTheoPhòngToolStripMenuItem_Click);
            // 
            // họcViênToolStripMenuItem
            // 
            this.họcViênToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.danhSáchHọcViênToolStripMenuItem,
            this.báoCáoĐiểmTheoLớpToolStripMenuItem,
            this.báoCáoCácHọcViênĐiểmCaoToolStripMenuItem});
            this.họcViênToolStripMenuItem.Name = "họcViênToolStripMenuItem";
            this.họcViênToolStripMenuItem.Size = new System.Drawing.Size(85, 24);
            this.họcViênToolStripMenuItem.Text = " Học viên";
            // 
            // danhSáchHọcViênToolStripMenuItem
            // 
            this.danhSáchHọcViênToolStripMenuItem.Name = "danhSáchHọcViênToolStripMenuItem";
            this.danhSáchHọcViênToolStripMenuItem.Size = new System.Drawing.Size(357, 26);
            this.danhSáchHọcViênToolStripMenuItem.Text = "Danh sách học viên";
            this.danhSáchHọcViênToolStripMenuItem.Click += new System.EventHandler(this.danhSáchHọcViênToolStripMenuItem_Click);
            // 
            // báoCáoĐiểmTheoLớpToolStripMenuItem
            // 
            this.báoCáoĐiểmTheoLớpToolStripMenuItem.Name = "báoCáoĐiểmTheoLớpToolStripMenuItem";
            this.báoCáoĐiểmTheoLớpToolStripMenuItem.Size = new System.Drawing.Size(357, 26);
            this.báoCáoĐiểmTheoLớpToolStripMenuItem.Text = "Báo cáo điểm theo lớp";
            this.báoCáoĐiểmTheoLớpToolStripMenuItem.Click += new System.EventHandler(this.báoCáoĐiểmTheoLớpToolStripMenuItem_Click);
            // 
            // báoCáoCácHọcViênĐiểmCaoToolStripMenuItem
            // 
            this.báoCáoCácHọcViênĐiểmCaoToolStripMenuItem.Name = "báoCáoCácHọcViênĐiểmCaoToolStripMenuItem";
            this.báoCáoCácHọcViênĐiểmCaoToolStripMenuItem.Size = new System.Drawing.Size(357, 26);
            this.báoCáoCácHọcViênĐiểmCaoToolStripMenuItem.Text = "Báo cáo các học viên điểm cao theo lớp";
            this.báoCáoCácHọcViênĐiểmCaoToolStripMenuItem.Click += new System.EventHandler(this.báoCáoCácHọcViênĐiểmCaoToolStripMenuItem_Click);
            // 
            // giáoViênToolStripMenuItem
            // 
            this.giáoViênToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.danhSáchGiáoViênToolStripMenuItem});
            this.giáoViênToolStripMenuItem.Name = "giáoViênToolStripMenuItem";
            this.giáoViênToolStripMenuItem.Size = new System.Drawing.Size(85, 24);
            this.giáoViênToolStripMenuItem.Text = "Giáo viên";
            // 
            // danhSáchGiáoViênToolStripMenuItem
            // 
            this.danhSáchGiáoViênToolStripMenuItem.Name = "danhSáchGiáoViênToolStripMenuItem";
            this.danhSáchGiáoViênToolStripMenuItem.Size = new System.Drawing.Size(225, 26);
            this.danhSáchGiáoViênToolStripMenuItem.Text = "Danh sách giáo viên";
            this.danhSáchGiáoViênToolStripMenuItem.Click += new System.EventHandler(this.danhSáchGiáoViênToolStripMenuItem_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::WindowsFormsApp3.Properties.Resources._5552862_image;
            this.pictureBox1.Location = new System.Drawing.Point(85, 101);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(624, 378);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 497);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Quản lý lớp học";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem lớpHọcToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem danhSáchLớpHọcToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem báoCáoDoanhThuTheoLớpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem báoCáoDanhSáchLớpTheoPhòngToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem họcViênToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem danhSáchHọcViênToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem báoCáoĐiểmTheoLớpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem báoCáoCácHọcViênĐiểmCaoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem giáoViênToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem danhSáchGiáoViênToolStripMenuItem;
    }
}

