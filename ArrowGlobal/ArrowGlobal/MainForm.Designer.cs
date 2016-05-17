namespace ArrowGlobal
{
    partial class MainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.sStrip = new System.Windows.Forms.StatusStrip();
            this.tsLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.cpbMain = new ArrowGlobal.CustomProgressBar();
            this.cpbSub = new ArrowGlobal.CustomProgressBar();
            this.dgView = new System.Windows.Forms.DataGridView();
            this.sStrip.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgView)).BeginInit();
            this.SuspendLayout();
            // 
            // sStrip
            // 
            this.sStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsLabel});
            this.sStrip.Location = new System.Drawing.Point(0, 649);
            this.sStrip.Name = "sStrip";
            this.sStrip.Size = new System.Drawing.Size(936, 22);
            this.sStrip.TabIndex = 0;
            this.sStrip.Text = "statusStrip1";
            // 
            // tsLabel
            // 
            this.tsLabel.Name = "tsLabel";
            this.tsLabel.Size = new System.Drawing.Size(39, 17);
            this.tsLabel.Text = "Ready";
            this.tsLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cpbMain
            // 
            this.cpbMain.CustomText = null;
            this.cpbMain.DisplayStyle = ArrowGlobal.ProgressBarDisplayText.Percentage;
            this.cpbMain.Location = new System.Drawing.Point(13, 13);
            this.cpbMain.Name = "cpbMain";
            this.cpbMain.Size = new System.Drawing.Size(358, 25);
            this.cpbMain.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.cpbMain.TabIndex = 1;
            // 
            // cpbSub
            // 
            this.cpbSub.CustomText = null;
            this.cpbSub.DisplayStyle = ArrowGlobal.ProgressBarDisplayText.Percentage;
            this.cpbSub.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.cpbSub.ForeColor = System.Drawing.Color.DeepSkyBlue;
            this.cpbSub.Location = new System.Drawing.Point(0, 634);
            this.cpbSub.Name = "cpbSub";
            this.cpbSub.Size = new System.Drawing.Size(936, 15);
            this.cpbSub.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.cpbSub.TabIndex = 2;
            // 
            // dgView
            // 
            this.dgView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgView.Location = new System.Drawing.Point(13, 45);
            this.dgView.Name = "dgView";
            this.dgView.Size = new System.Drawing.Size(911, 583);
            this.dgView.TabIndex = 3;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(936, 671);
            this.Controls.Add(this.dgView);
            this.Controls.Add(this.cpbSub);
            this.Controls.Add(this.cpbMain);
            this.Controls.Add(this.sStrip);
            this.DoubleBuffered = true;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "MainForm";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.sStrip.ResumeLayout(false);
            this.sStrip.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.StatusStrip sStrip;
        public CustomProgressBar cpbMain;
        public CustomProgressBar cpbSub;
        public System.Windows.Forms.ToolStripStatusLabel tsLabel;
        private System.Windows.Forms.DataGridView dgView;
    }
}

