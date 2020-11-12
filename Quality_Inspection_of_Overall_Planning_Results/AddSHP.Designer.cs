namespace Quality_Inspection_of_Overall_Planning_Results
{
    partial class AddSHP
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
            this.lb_main = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btn_ok = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.lb_new = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // lb_main
            // 
            this.lb_main.FormattingEnabled = true;
            this.lb_main.ItemHeight = 12;
            this.lb_main.Location = new System.Drawing.Point(9, 22);
            this.lb_main.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.lb_main.Name = "lb_main";
            this.lb_main.Size = new System.Drawing.Size(294, 76);
            this.lb_main.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 5);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "选择主图层";
            // 
            // btn_ok
            // 
            this.btn_ok.Location = new System.Drawing.Point(117, 206);
            this.btn_ok.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btn_ok.Name = "btn_ok";
            this.btn_ok.Size = new System.Drawing.Size(62, 24);
            this.btn_ok.TabIndex = 3;
            this.btn_ok.Text = "合并";
            this.btn_ok.UseVisualStyleBackColor = true;
            this.btn_ok.Click += new System.EventHandler(this.btn_ok_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(10, 111);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "新增图层";
            // 
            // lb_new
            // 
            this.lb_new.FormattingEnabled = true;
            this.lb_new.ItemHeight = 12;
            this.lb_new.Location = new System.Drawing.Point(9, 126);
            this.lb_new.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.lb_new.Name = "lb_new";
            this.lb_new.Size = new System.Drawing.Size(294, 76);
            this.lb_new.TabIndex = 5;
            // 
            // AddSHP
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(311, 235);
            this.Controls.Add(this.lb_new);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btn_ok);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lb_main);
            this.Name = "AddSHP";
            this.Text = "合并图层";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox lb_main;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn_ok;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ListBox lb_new;
    }
}