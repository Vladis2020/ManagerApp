using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;

using System.Windows.Forms;

namespace Reg
{
    public partial class ChooseTable : Form
    {
        public ChooseTable()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var f2 = new Form2(this);
            this.Visible = false;
            if (comboBox1.SelectedIndex == 0)
            {
                f2.Text = "Договор";

            }
            else
              if (comboBox1.SelectedIndex == 1)
            {
                f2.Text = "Клиент";
            }
            else
              if (comboBox1.SelectedIndex == 2)
            {
                f2.Text = "Менеджер";
            }
            else
              if (comboBox1.SelectedIndex == 3)
            {
                f2.Text = "Отчет";
            }
            else
              if (comboBox1.SelectedIndex == 4)
            {
                f2.Text = "Переговоры";
            }
            else
              if (comboBox1.SelectedIndex == 5)
            {
                f2.Text = "Услуги";
            }

            f2.Show();

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ChooseTable_Load(object sender, EventArgs e)
        {

        }

        private void ChooseTable_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        private void ChooseTable_FormClosed(object sender, FormClosedEventArgs e)
        {

            Application.Exit();
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }

}