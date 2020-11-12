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
    public partial class SelectForm : Form
    {
        public SelectForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            frmMain frm = new frmMain();
            frm.Show();
        }

        private void SelectForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult result = MessageBox.Show("将关闭所有任务，确认退出吗?", "操作提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (result == DialogResult.OK)
            {
                Dispose();
                Application.Exit();
            }
            else
            {
                e.Cancel = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            frmMainQuZong frm = new frmMainQuZong();
            frm.Show();
        }

        private void pictureBox1_MouseHover(object sender, EventArgs e)
        {
            pictureBox1.BackgroundImage = null;
        }

        private void pictureBox1_MouseLeave(object sender, EventArgs e)
        {
            pictureBox1.BackgroundImage = Image.FromFile(Application.StartupPath + @"\button1.png");
        }

        private void pictureBox2_MouseHover(object sender, EventArgs e)
        {
            pictureBox2.BackgroundImage = null;
        }

        private void pictureBox3_MouseHover(object sender, EventArgs e)
        {
            pictureBox3.BackgroundImage = null;
        }

        private void pictureBox4_MouseHover(object sender, EventArgs e)
        {
            pictureBox4.BackgroundImage = null;
        }

        private void pictureBox5_MouseHover(object sender, EventArgs e)
        {
            pictureBox5.BackgroundImage = null;
        }

        private void pictureBox6_MouseHover(object sender, EventArgs e)
        {
            pictureBox6.BackgroundImage = null;
        }

        private void pictureBox7_MouseHover(object sender, EventArgs e)
        {
            pictureBox7.BackgroundImage = null;
        }

        private void pictureBox2_MouseLeave(object sender, EventArgs e)
        {
            pictureBox2.BackgroundImage = Image.FromFile(Application.StartupPath + @"\button2.png");
        }

        private void pictureBox3_MouseLeave(object sender, EventArgs e)
        {
            pictureBox3.BackgroundImage = Image.FromFile(Application.StartupPath + @"\button3.png");
        }

        private void pictureBox4_MouseLeave(object sender, EventArgs e)
        {
            pictureBox4.BackgroundImage = Image.FromFile(Application.StartupPath + @"\button4.png");
        }

        private void pictureBox5_MouseLeave(object sender, EventArgs e)
        {
            pictureBox5.BackgroundImage = Image.FromFile(Application.StartupPath + @"\button5.png");
        }

        private void pictureBox6_MouseLeave(object sender, EventArgs e)
        {
            pictureBox6.BackgroundImage = Image.FromFile(Application.StartupPath + @"\button6.png");
        }

        private void pictureBox7_MouseLeave(object sender, EventArgs e)
        {
            pictureBox7.BackgroundImage = Image.FromFile(Application.StartupPath + @"\button7.png");
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            frmMain2020 frm = new frmMain2020();
            frm.Show();
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            ZenJian frm = new ZenJian();
            frm.Show();
        }
    }
}
