using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;


namespace Reg
{
    public partial class Form3 : Form
    {

        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Максим\Downloads\Reg\Reg\Курсовая бд.accdb");

        public Form3()
        {
            InitializeComponent();
        }


       


      

        private void button1_Click(object sender, EventArgs e)
        {
            Form2 form = (Form2)this.Owner;
            



            form.Show();
            this.Close();


        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
