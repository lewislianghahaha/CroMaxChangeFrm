namespace CroMaxChangeFrm
{
    partial class Main
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.panel1 = new System.Windows.Forms.Panel();
            this.rbFormualChange = new System.Windows.Forms.RadioButton();
            this.rbColorantForChange = new System.Windows.Forms.RadioButton();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btnopen = new System.Windows.Forms.Button();
            this.MainMenu = new System.Windows.Forms.MenuStrip();
            this.tmclose = new System.Windows.Forms.ToolStripMenuItem();
            this.btngen = new System.Windows.Forms.Button();
            this.btnexport = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.MainMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.rbColorantForChange);
            this.panel1.Controls.Add(this.rbFormualChange);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 25);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(243, 53);
            this.panel1.TabIndex = 0;
            // 
            // rbFormualChange
            // 
            this.rbFormualChange.AutoSize = true;
            this.rbFormualChange.Location = new System.Drawing.Point(13, 8);
            this.rbFormualChange.Name = "rbFormualChange";
            this.rbFormualChange.Size = new System.Drawing.Size(71, 16);
            this.rbFormualChange.TabIndex = 0;
            this.rbFormualChange.TabStop = true;
            this.rbFormualChange.Text = "格式转换";
            this.rbFormualChange.UseVisualStyleBackColor = true;
            // 
            // rbColorantForChange
            // 
            this.rbColorantForChange.AutoSize = true;
            this.rbColorantForChange.Location = new System.Drawing.Point(13, 31);
            this.rbColorantForChange.Name = "rbColorantForChange";
            this.rbColorantForChange.Size = new System.Drawing.Size(119, 16);
            this.rbColorantForChange.TabIndex = 1;
            this.rbColorantForChange.TabStop = true;
            this.rbColorantForChange.Text = "色母相关格式转换";
            this.rbColorantForChange.UseVisualStyleBackColor = true;
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel2.Controls.Add(this.btnexport);
            this.panel2.Controls.Add(this.btngen);
            this.panel2.Controls.Add(this.btnopen);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 78);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(243, 122);
            this.panel2.TabIndex = 1;
            // 
            // btnopen
            // 
            this.btnopen.Location = new System.Drawing.Point(27, 17);
            this.btnopen.Name = "btnopen";
            this.btnopen.Size = new System.Drawing.Size(191, 23);
            this.btnopen.TabIndex = 0;
            this.btnopen.Text = "导入EXCEL";
            this.btnopen.UseVisualStyleBackColor = true;
            // 
            // MainMenu
            // 
            this.MainMenu.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.MainMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tmclose});
            this.MainMenu.Location = new System.Drawing.Point(0, 0);
            this.MainMenu.Name = "MainMenu";
            this.MainMenu.Size = new System.Drawing.Size(243, 25);
            this.MainMenu.TabIndex = 2;
            this.MainMenu.Text = "MainMenu";
            // 
            // tmclose
            // 
            this.tmclose.Name = "tmclose";
            this.tmclose.Size = new System.Drawing.Size(44, 21);
            this.tmclose.Text = "关闭";
            // 
            // btngen
            // 
            this.btngen.Location = new System.Drawing.Point(27, 46);
            this.btngen.Name = "btngen";
            this.btngen.Size = new System.Drawing.Size(191, 23);
            this.btngen.TabIndex = 1;
            this.btngen.Text = "运算";
            this.btngen.UseVisualStyleBackColor = true;
            // 
            // btnexport
            // 
            this.btnexport.Location = new System.Drawing.Point(27, 75);
            this.btnexport.Name = "btnexport";
            this.btnexport.Size = new System.Drawing.Size(191, 23);
            this.btnexport.TabIndex = 2;
            this.btnexport.Text = "导出EXCEL";
            this.btnexport.UseVisualStyleBackColor = true;
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(243, 200);
            this.ControlBox = false;
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.MainMenu);
            this.MainMenuStrip = this.MainMenu;
            this.Name = "Main";
            this.Text = "科丽晶数据转换";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.MainMenu.ResumeLayout(false);
            this.MainMenu.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.RadioButton rbColorantForChange;
        private System.Windows.Forms.RadioButton rbFormualChange;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button btnopen;
        private System.Windows.Forms.MenuStrip MainMenu;
        private System.Windows.Forms.Button btnexport;
        private System.Windows.Forms.Button btngen;
        private System.Windows.Forms.ToolStripMenuItem tmclose;
    }
}

