using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Quality_Inspection_of_Overall_Planning_Results
{
    public partial class Change : Form
    {
        string fieldname = "";
        string fieldvalue = "";
        public event EventHandler<ChangeEventArgs> ChangeOK;

        public Change(string _fieldname)
        {
            InitializeComponent();
            fieldname = _fieldname;
            label1.Text = _fieldname;
        }

        public void button1_Click(object sender, EventArgs e)
        {
            if (this.textBox1.Text == "" || this.textBox1.Text == null) { }
            else { fieldvalue = textBox1.Text; }

            if (fieldvalue == "")
            {
                base.Close();
            }

            if (this.ChangeOK != null)
                this.ChangeOK(this, new ChangeEventArgs()
                {
                    field_value = this.fieldvalue,
                });

            base.Close();
        }
    }
}
