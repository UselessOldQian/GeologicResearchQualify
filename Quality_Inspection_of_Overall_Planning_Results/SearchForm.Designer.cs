namespace Quality_Inspection_of_Overall_Planning_Results
{
    partial class SearchForm
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
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.ListFieldsName = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.comLayerName = new System.Windows.Forms.ComboBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 12;
            this.listBox1.Location = new System.Drawing.Point(12, 222);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(535, 112);
            this.listBox1.TabIndex = 0;
            this.listBox1.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.listbox1_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(81, 354);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 2;
            this.btnOK.Text = "查询";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.button1_Click);
            // 
            // ListFieldsName
            // 
            this.ListFieldsName.FormattingEnabled = true;
            this.ListFieldsName.Location = new System.Drawing.Point(65, 51);
            this.ListFieldsName.Name = "ListFieldsName";
            this.ListFieldsName.Size = new System.Drawing.Size(249, 20);
            this.ListFieldsName.TabIndex = 5;
            this.ListFieldsName.SelectedIndexChanged += new System.EventHandler(this.ListFieldsName_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(330, 57);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(11, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "=";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(18, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "图层：";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(365, 50);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(204, 21);
            this.textBox1.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(18, 54);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 12);
            this.label1.TabIndex = 3;
            this.label1.Text = "属性：";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(585, 49);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(46, 23);
            this.button1.TabIndex = 8;
            this.button1.Text = "并且";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // comLayerName
            // 
            this.comLayerName.FormattingEnabled = true;
            this.comLayerName.Location = new System.Drawing.Point(65, 17);
            this.comLayerName.Name = "comLayerName";
            this.comLayerName.Size = new System.Drawing.Size(249, 20);
            this.comLayerName.TabIndex = 1;
            this.comLayerName.SelectedIndexChanged += new System.EventHandler(this.comLayerName_SelectedIndexChanged);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.comLayerName);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.textBox1);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.ListFieldsName);
            this.panel1.Location = new System.Drawing.Point(16, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(667, 213);
            this.panel1.TabIndex = 9;
            // 
            // SearchForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(719, 389);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.listBox1);
            this.Name = "SearchForm";
            this.Text = "SearchForm";
            this.TopMost = true;
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.ComboBox ListFieldsName;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ComboBox comLayerName;
        private System.Windows.Forms.Panel panel1;
    }
}