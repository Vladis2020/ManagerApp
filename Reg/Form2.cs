using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Text.RegularExpressions;

namespace Reg
{
    public partial class Form2 : Form
    {
        string ID = string.Empty;
        int selected = -1;
        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\menager.accdb;");
        private readonly ChooseTable chooseTable;
        public Form2(ChooseTable chooseTable)
        {
            InitializeComponent();
            this.chooseTable = chooseTable;
        }
        private void loadDGView(string query)
        {
           
            con.Open();
            OleDbDataAdapter adapter = new OleDbDataAdapter(query, con);
            DataTable dataTable = new DataTable();
            adapter.Fill(dataTable);
            con.Close();

            dataGridView1.DataSource = dataTable;
            dataGridView1.Refresh();
        }
        
        public void LoadComba()
        {
            BindingList<CbbEntry> cbbEntries = new BindingList<CbbEntry>();
            BindingList<CbbEntry> cbbEntries1 = new BindingList<CbbEntry>();
            BindingList<CbbEntry> cbbEntries2 = new BindingList<CbbEntry>();
            try
            {
                con.Open();
                var cmd = con.CreateCommand();
                var cmd2 = con.CreateCommand();
                var cmd3 = con.CreateCommand();
                switch (this.Text)
                {
                    case "Договор":
                        {
                            cmd.CommandText = "select КодКлиента, НазваниеКомпании from Клиент";
                            OleDbDataReader reader = cmd.ExecuteReader();
                            while (reader.Read())
                            {
                                cbbEntries.Add(new CbbEntry((int)reader["КодКлиента"], (string)reader["НазваниеКомпании"]));
                            }
                            cmd2.CommandText = "select КодУслуги, ВидУслуги from Услуги";
                            OleDbDataReader reader2 = cmd2.ExecuteReader();
                            while (reader2.Read())
                            {
                                cbbEntries1.Add(new CbbEntry((int)reader2[0], (string)reader2[1]));
                            }
                            cmd3.CommandText = "select КодМенеджера, Фамилия from Менеджер";
                            OleDbDataReader reader3 = cmd3.ExecuteReader();
                            while (reader3.Read())
                            {
                                cbbEntries2.Add(new CbbEntry((int)reader3[0], (string)reader3[1]));
                            }
                        }
                        break;
                    case "Клиент":
                        {
                            cmd.CommandText = "select КодУслуги, ВидУслуги from Услуги";
                            OleDbDataReader reader = cmd.ExecuteReader();
                            while (reader.Read())
                            {
                                cbbEntries.Add(new CbbEntry((int)reader["КодУслуги"], (string)reader["ВидУслуги"]));
                            }
                        }
                        break;
                    case "Отчет":
                        {
                            cmd.CommandText = "select КодМенеджера, фамилия from Менеджер";
                            OleDbDataReader reader = cmd.ExecuteReader();
                            while (reader.Read())
                            {
                                cbbEntries.Add(new CbbEntry((int)reader["КодФакультета"], (string)reader["название"]));
                            }
                            
                        }
                        break;
                    case "Переговоры":
                        {
                            cmd.CommandText = "select КодУслуги, ВидУслуги from Услуги";
                            OleDbDataReader reader = cmd.ExecuteReader();
                            while (reader.Read())
                            {
                                cbbEntries.Add(new CbbEntry((int)reader["КодУслуги"], (string)reader["ВидУслуги"]));
                            }
                            cmd2.CommandText = "select КодКлиента, НазваниеКомпании from Клиент";
                            OleDbDataReader reader2 = cmd2.ExecuteReader();
                            while (reader2.Read())
                            {
                                cbbEntries1.Add(new CbbEntry((int)reader2[0], (string)reader2[1]));
                            }
                            cmd3.CommandText = "select КодМенеджера, Фамилия from Менеджер";
                            OleDbDataReader reader3 = cmd3.ExecuteReader();
                            while (reader3.Read())
                            {
                                cbbEntries2.Add(new CbbEntry((int)reader3[0], (string)reader3[1]));
                            }

                        }
                        break;
                }
                comboBox1.DataSource = cbbEntries;
                comboBox2.DataSource = cbbEntries1;
                comboBox3.DataSource = cbbEntries2;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            con.Close();
        }
        private void Form2_Shown(object sender, EventArgs e)
        {
            поискToolStripMenuItem.Visible = Text == "Услуги" || Text == "Клиент" || Text == "Менеджер";

            if (this.Text == "Услуги")
            {
                loadDGView(Queris.selectUslugi);
                dataGridView1.Columns[0].Visible = false;
            }
            else
            if (this.Text == "Клиент")
            {
                LoadComba();
                loadDGView(Queris.selectKlient);
                dataGridView1.Columns[0].Visible = false;
                dataGridView1.Columns["КодУслуги"].Visible = false;
                button2.Visible = button3.Visible = button4.Visible = true;
            }

            else
             if (this.Text == "Отчет")
            {
                loadDGView(Queris.selectOtchet);
                dataGridView1.Columns[0].Visible = false;
                dataGridView1.Columns["КодМенеджера"].Visible = false;
                LoadComba();
            }
            else
             if (this.Text == "Договор")
            {
                loadDGView(Queris.selectDogovor);
                dataGridView1.Columns[0].Visible = false;
                dataGridView1.Columns["КодКлиента"].Visible = false;
                dataGridView1.Columns["КодУслуги"].Visible = false;
                dataGridView1.Columns["КодМенеджера"].Visible = false;
                LoadComba();
            }
            else
             if (this.Text == "Менеджер")
            {
                loadDGView(Queris.selectMeneger);
                dataGridView1.Columns[0].Visible = false;
            }
            else
             if (this.Text == "Переговоры")
            {
                loadDGView(Queris.selectPeregovory);
                dataGridView1.Columns[0].Visible = false;
                dataGridView1.Columns["КодУслуги"].Visible = false;
                dataGridView1.Columns["КодКлиента"].Visible = false;
                dataGridView1.Columns["КодМенеджера"].Visible = false;
                LoadComba();
            }
        }

        private void добавлениеToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            if (this.Text == "Договор")
            {
                textBox1.Visible = true;
                comboBox1.Visible = true;
                comboBox2.Visible = true;
                comboBox3.Visible = true;
                label6.Visible = true;
                label7.Visible = true;
                label8.Visible = true;
                label1.Visible = true;
                label1.Text = "дата";
                label6.Text = "клиент";
                label7.Text = "услуга";
                label8.Text = "менеджер";
                panel1.Visible = true;
                button1.Visible = true;
                button1.BackColor = Color.Yellow;
                panel1.BackColor = Color.Snow;
            }
            else
            if (this.Text == "Переговоры")
            {
                textBox1.Visible = true;
                textBox2.Visible = true;
               
                comboBox1.Visible = true;
                comboBox2.Visible = true;
                comboBox3.Visible = true;
                label6.Visible = true;
                label7.Visible = true;
                label8.Visible = true;
                label1.Visible = true;
                label2.Visible = true;
                
                label6.Text = "Услуги";
                label7.Text = "Клиент";
                label8.Text = "Менеджер";
                label1.Text = "Контактный адрес";
                label2.Text = "статус";
                button1.Visible = true;
                button1.BackColor = Color.Yellow;
                panel1.BackColor = Color.Snow;

            }
            else
            if (this.Text == "Клиент")
            {
                textBox1.Visible = true;
                label6.Visible = true;
                label1.Visible = true;
                comboBox1.Visible = true;
               
                label1.Text = "Название Компании";
                label6.Text = "Услуга";
               
                button1.Visible = true;
                button1.BackColor = Color.Yellow;
                panel1.BackColor = Color.Snow;
            }
            else
            if (this.Text == "Менеджер")
            {
                textBox1.Visible = true;
                textBox2.Visible = true;
                textBox3.Visible = true;
                textBox4.Visible = true;
                textBox5.Visible = true;
                label1.Visible = true;
                label2.Visible = true;
                label3.Visible = true;
                label4.Visible = true;
                label5.Visible = true;
                label1.Text = "Фамилия";
                label2.Text = "Имя";
                label3.Text = "Отчество";
                label4.Text = "Компания";
                label5.Text = "Должность";

                button1.Visible = true;
                button1.BackColor = Color.Yellow;
                panel1.BackColor = Color.Snow;
            }
            if (Text == "Услуги")
            {
                label1.Text = "ВидУслуги";
                label1.Visible = true;
                textBox1.Visible = true;
                button1.Visible = true;
                button1.BackColor = Color.Yellow;
                panel1.BackColor = Color.Snow;
            }
            else
            if (this.Text == "Отчет")
            {
                textBox1.Visible = true;
                label6.Visible = true;
                label1.Visible = true;
                comboBox1.Visible = true;

                label1.Text = "Количество Заключенных Договоров";
                label6.Text = "Менеджер";

                button1.Visible = true;
                button1.BackColor = Color.Yellow;
                panel1.BackColor = Color.Snow;
            }

        }
        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (InputCorrect())
            {
                if ((this.Text == "Услуги") && (button1.BackColor == Color.Yellow))
                {
                    con.Open();
                    string query = "INSERT INTO Услуги(ВидУслуги) VALUES (@p1)";
                    OleDbCommand command = new OleDbCommand(query, con);
                    command.Parameters.AddWithValue("@p1", textBox1.Text);
                   
                    command.ExecuteNonQuery();
                    con.Close();
                    loadDGView(Queris.selectUslugi);
                }
                else
                if ((this.Text == "Услуги") && (button1.BackColor == Color.Snow))
                {
                    Izmen();
                }
                else
                if ((this.Text == "Договор") && (button1.BackColor == Color.Yellow))
                {
                    con.Open();
                    string query = "INSERT INTO Договор(КодКлиента, КодУслуги, КодМенеджера, Дата) VALUES (@p1,@p2,@p3,@p4)";
                    OleDbCommand command = new OleDbCommand(query, con);
                    command.Parameters.AddWithValue("@p1", comboBox1.SelectedValue);
                    command.Parameters.AddWithValue("@p2", comboBox2.SelectedValue);
                    command.Parameters.AddWithValue("@p3", comboBox3.SelectedValue);
                    command.Parameters.AddWithValue("@p4", textBox1.Text);
                    
                    command.ExecuteNonQuery();
                    con.Close();
                    loadDGView(Queris.selectDogovor);
                }
                else
                if ((this.Text == "Договор") && (button1.BackColor == Color.Snow))
                {
                    Izmen();
                }
                else
                if ((this.Text == "Переговоры") && (button1.BackColor == Color.Yellow))
                {
                    con.Open();
                    string query = "INSERT INTO Переговоры(КодУслуги, КодКлиента, КодМенеджера, КонтактныйАдрес, Статус) VALUES (@p1,@p2,@p3,@p4,@p5)";
                    OleDbCommand command = new OleDbCommand(query, con);
                    command.Parameters.AddWithValue("@p1", comboBox1.SelectedValue);
                    command.Parameters.AddWithValue("@p2", comboBox2.SelectedValue);
                    command.Parameters.AddWithValue("@p3", comboBox3.SelectedValue);
                    command.Parameters.AddWithValue("@p4", textBox1.Text);
                    command.Parameters.AddWithValue("@p5", textBox2.Text);
                    command.ExecuteNonQuery();
                    con.Close();
                    loadDGView(Queris.selectPeregovory);
                }
                else
                if ((this.Text == "Переговоры") && (button1.BackColor == Color.Snow))
                {
                    Izmen();
                }
                else
                if ((this.Text == "Клиент") && (button1.BackColor == Color.Yellow))
                {
                    con.Open();
                    string query = "INSERT INTO Клиент(НазваниеКомпании, КодУслуги) VALUES (@p1,@p2)";
                    OleDbCommand command = new OleDbCommand(query, con);
                    command.Parameters.AddWithValue("@p1", textBox1.Text);
                    command.Parameters.AddWithValue("@p2", comboBox1.SelectedValue);
                    command.ExecuteNonQuery();
                    con.Close();

                    loadDGView(Queris.selectKlient);
                }
                else
                if ((this.Text == "Клиент") && (button1.BackColor == Color.Snow))
                {
                    Izmen();
                }
                if ((this.Text == "Отчет") && (button1.BackColor == Color.Yellow))
                {
                    con.Open();
                    string query = "INSERT INTO Отчет(КодМенеджера, КоличествоЗаключенныхДоговоров) VALUES (@p1,@p2)";
                    OleDbCommand command = new OleDbCommand(query, con);
                    command.Parameters.AddWithValue("@p1", comboBox1.SelectedValue);
                    command.Parameters.AddWithValue("@p2", textBox1.Text);
                    command.ExecuteNonQuery();
                    con.Close();
                    loadDGView(Queris.selectOtchet);
                }
                else
                if ((this.Text == "Отчет") && (button1.BackColor == Color.Snow))
                {
                    Izmen();
                }
                else
                if ((this.Text == "Менеджер") && (button1.BackColor == Color.Yellow))
                {
                    con.Open();
                    string query = "INSERT INTO Менеджер(Фамилия, Имя, Отчество, Должность, Компания) VALUES (@p1,@p2,@p3,@p4,@p5)";
                    OleDbCommand command = new OleDbCommand(query, con);
                    command.Parameters.AddWithValue("@p1", textBox1.Text);
                    command.Parameters.AddWithValue("@p2", textBox2.Text);
                    command.Parameters.AddWithValue("@p3", textBox3.Text);
                    command.Parameters.AddWithValue("@p4", textBox4.Text);
                    command.Parameters.AddWithValue("@p5", textBox5.Text);
                    command.ExecuteNonQuery();
                    con.Close();
                    loadDGView(Queris.selectMeneger);
                }
                else
                if ((this.Text == "Менеджер") && (button1.BackColor == Color.Snow))
                {
                    Izmen();
                }
            }
            else
            {
                MessageBox.Show("Вы ввели некорректные данные!","Ошибка!");
            }
        }

        private bool InputCorrect()
        {
            const string буквы = @"^[А-я]+$";
            const string буквы_и_знаки_препинания = @"^[А-я\-., ]+$";
            const string буквы_цифры_знаки_препинания = @"^[А-я\-.,0-9 ]+$";
            const string цифры = @"^\d+$";
            switch (Text)
            {
                case "Услуги":
                    return Regex.IsMatch(textBox1.Text, буквы_и_знаки_препинания);
                           
                case "Договор":
                    return  
                            Regex.IsMatch(textBox1.Text, буквы_цифры_знаки_препинания) &&
                            comboBox1.SelectedItem != null &&
                            comboBox2.SelectedItem != null &&
                            comboBox3.SelectedItem != null;
                case "Клиент":
                    return  Regex.IsMatch(textBox1.Text, буквы_и_знаки_препинания) &&
                            
                           
                            comboBox1.SelectedItem != null;
                case "Отчет":
                    return comboBox1.SelectedItem != null &&
                           Regex.IsMatch(textBox1.Text, цифры);
                case "Переговоры":
                    return
                    Regex.IsMatch(textBox1.Text, буквы_цифры_знаки_препинания) &&
                    Regex.IsMatch(textBox2.Text, буквы) &&
                            comboBox1.SelectedItem != null &&
                            comboBox2.SelectedItem != null &&
                            comboBox3.SelectedItem != null;
                case "Менеджер":
                    return
                    Regex.IsMatch(textBox1.Text, буквы) &&
                    Regex.IsMatch(textBox2.Text, буквы) &&
                    Regex.IsMatch(textBox2.Text, буквы) &&
                    Regex.IsMatch(textBox2.Text, буквы) &&
                    Regex.IsMatch(textBox2.Text, буквы_цифры_знаки_препинания);

                default:
                    return false;
            }
        }

        private void изменениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.Text == "Отчет")
                {
                    comboBox1.SelectedValue = int.Parse(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells["КодМенеджера"].Value.ToString());
                    textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells["КоличествоЗаключенныхДоговоров"].Value.ToString();
                    textBox1.Visible = true;
                    label6.Visible = true;
                    label1.Visible = true;
                    comboBox1.Visible = true;

                    label1.Text = "Количество Заключенных Договоров";
                    label6.Text = "Менеджер";

                    panel1.Visible = true;
                    button1.Visible = true;
                    button1.BackColor = Color.Snow;
                    panel1.BackColor = Color.Snow;
                }
                else
            if (this.Text == "Переговоры")
                {
                    textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells["КонтактныйАдрес"].Value.ToString();
                    textBox2.Text = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells["Статус"].Value.ToString();
                  
                    comboBox1.SelectedValue = int.Parse(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells["КодУслуги"].Value.ToString());
                    comboBox2.SelectedValue = int.Parse(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells["КодКлиента"].Value.ToString());
                    comboBox3.SelectedValue = int.Parse(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells["КодМенеджера"].Value.ToString());
                    textBox1.Visible = true;
                    textBox2.Visible = true;

                    comboBox1.Visible = true;
                    comboBox2.Visible = true;
                    comboBox3.Visible = true;
                    label6.Visible = true;
                    label7.Visible = true;
                    label8.Visible = true;
                    label1.Visible = true;
                    label2.Visible = true;

                    label6.Text = "Услуги";
                    label7.Text = "Клиент";
                    label8.Text = "Менеджер";
                    label1.Text = "Контактный адрес";
                    label2.Text = "статус";
                    button1.Visible = true;
                    
                    button1.BackColor = Color.Snow;
                    panel1.BackColor = Color.Snow;

                }
                else
            if (this.Text == "Клиент")
                {
                    textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells["НазваниеКомпании"].Value.ToString();
                   
                    comboBox1.SelectedValue = int.Parse(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells["КодУслуги"].Value.ToString());
                    textBox1.Visible = true;
                    label6.Visible = true;
                    label1.Visible = true;
                    comboBox1.Visible = true;

                    label1.Text = "Название Компании";
                    label6.Text = "Услуга";
                    button1.Visible = true;
                    button1.BackColor = Color.Snow;
                    panel1.BackColor = Color.Snow;
                }
                else
            if (this.Text == "Договор")
                {
                    textBox1.Text = ((DateTime)dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells["Дата"].Value).ToShortDateString().ToString();
                    comboBox1.SelectedValue = int.Parse(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells["КодКлиента"].Value.ToString());
                    comboBox2.SelectedValue = int.Parse(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells["КодУслуги"].Value.ToString());
                    comboBox3.SelectedValue = int.Parse(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells["КодМенеджера"].Value.ToString());
                    textBox1.Visible = true;
                    comboBox1.Visible = true;
                    comboBox2.Visible = true;
                    comboBox3.Visible = true;
                    label6.Visible = true;
                    label7.Visible = true;
                    label8.Visible = true;
                    label1.Visible = true;
                    label2.Visible = true;
                    label1.Text = "дата";
                    label6.Text = "клиент";
                    label7.Text = "менеджер";
                    label8.Text = "услуга";
                    button1.Visible = true;
                    panel1.Visible = true;
                    button1.BackColor = Color.Snow;
                    panel1.BackColor = Color.Snow;
                }
                else
            if (this.Text == "Услуги")
                {
                    textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells["ВидУслуги"].Value.ToString();
                    label1.Text = "ВидУслуги";
                    label1.Visible = true;
                    textBox1.Visible = true;
                    button1.Visible = true;
                    button1.BackColor = Color.Snow;
                    panel1.BackColor = Color.Snow;
                }
                else
            if (this.Text == "Менеджер")
                {
                    textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells["Фамилия"].Value.ToString();
                    textBox2.Text = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells["Имя"].Value.ToString();
                    textBox3.Text = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells["Отчество"].Value.ToString();
                    textBox4.Text = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells["Компания"].Value.ToString();
                    textBox5.Text = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells["Должность"].Value.ToString();
                    textBox1.Visible = true;
                    textBox2.Visible = true;
                    textBox3.Visible = true;
                    textBox4.Visible = true;
                    textBox5.Visible = true;
                    label1.Visible = true;
                    label2.Visible = true;
                    label3.Visible = true;
                    label4.Visible = true;
                    label5.Visible = true;
                    label1.Text = "Фамилия";
                    label2.Text = "Имя";
                    label3.Text = "Отчество";
                    label4.Text = "Компания";
                    label5.Text = "Должность";
                    button1.Visible = true;
                    button1.BackColor = Color.Snow;
                    panel1.BackColor = Color.Snow;
                }
            }
            catch (Exception )
            {
                MessageBox.Show("Вы не выбрали поля для изменения");
            }

        }

        private void Izmen()
        {
            if (MessageBox.Show("Уверенны что хотите сохранить эту запись?", "Сохранить?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                switch (Text)
                {
                    case "Договор":
                        {
                            int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString());
                            con.Open();
                            string query = "UPDATE Договор SET КодКлиента=@p1, КодУслуги=@p2, КодМенеджера=@p3, Дата=@p4 WHERE КодДоговора=@id";
                            OleDbCommand command = new OleDbCommand(query, con);
                            command.Parameters.AddWithValue("@p1", comboBox1.SelectedValue);
                            command.Parameters.AddWithValue("@p2", comboBox2.SelectedValue);
                            command.Parameters.AddWithValue("@p3", comboBox3.SelectedValue);
                            command.Parameters.Add("@p4", OleDbType.Date).Value = DateTime.Parse(textBox1.Text);
                            
                            command.Parameters.AddWithValue("@id", ID);
                            command.ExecuteNonQuery();
                            con.Close();
                            loadDGView(Queris.selectDogovor);
                        }
                        break;
                    case "Переговоры":
                        {
                            int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString());
                            con.Open();
                            string query = "UPDATE Переговоры SET КодУслуги=@p1, КодКлиента=@p2, КодМенеджера=@p3, КонтактныйАдрес=@p4, Статус=@p5 WHERE КодПереговоров=@id";
                            OleDbCommand command = new OleDbCommand(query, con);
                            command.Parameters.AddWithValue("@p1", comboBox1.SelectedValue);
                            command.Parameters.AddWithValue("@p2", comboBox2.SelectedValue);
                            command.Parameters.AddWithValue("@p3", comboBox3.SelectedValue);
                            command.Parameters.AddWithValue("@p4", textBox1.Text);
                            command.Parameters.AddWithValue("@p5", textBox2.Text);
                            command.Parameters.AddWithValue("@id", ID);
                            command.ExecuteNonQuery();
                            con.Close();
                            loadDGView(Queris.selectPeregovory);
                        }
                        break;
                    case "Отчет":
                        {
                            int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString());
                            con.Open();
                            string query = "UPDATE Отчет SET КодМенеджера=@p1, КоличествоЗаключенныхДоговоров=@p2 WHERE КодОтчета=@id";
                            OleDbCommand command = new OleDbCommand(query, con);
                            command.Parameters.AddWithValue("@p1", comboBox1.SelectedValue);
                            command.Parameters.AddWithValue("@p2", textBox1.Text);
                            
                           
                            command.Parameters.AddWithValue("@id", ID);
                            command.ExecuteNonQuery();
                            con.Close();
                            loadDGView(Queris.selectOtchet);
                        }
                        break;
                    case "Клиент":
                        {
                            int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString());
                            con.Open();
                            string query = "UPDATE Клиент SET НазваниеКомпании=@p1, КодУслуги=@p2 WHERE КодКлиента=@id";
                            OleDbCommand command = new OleDbCommand(query, con);
                            command.Parameters.AddWithValue("@p4", textBox1.Text);
                            command.Parameters.AddWithValue("@p2", comboBox1.SelectedValue);
                            command.Parameters.AddWithValue("@id", ID);
                            command.ExecuteNonQuery();
                            con.Close();
                            loadDGView(Queris.selectKlient);
                        }
                        break;
                    case "Услуги":
                        {
                            int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString());
                            con.Open();
                            string query = "UPDATE Услуги SET ВидУслуги=@p1 WHERE КодУслуги=@id";
                            OleDbCommand command = new OleDbCommand(query, con);
                            command.Parameters.AddWithValue("@p1", textBox1.Text);
                            command.Parameters.AddWithValue("@id", ID);
                            command.ExecuteNonQuery();
                            con.Close();
                            loadDGView(Queris.selectUslugi);
                        }
                        break;
                    case "Менеджер":
                        {
                            int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString());
                            con.Open();
                            string query = "UPDATE Менеджер SET Фамилия=@p1, Имя=@p2, Отчество=@p3, Компания=@p4, Должность=@p5 WHERE КодМенеджера=@id";
                            OleDbCommand command = new OleDbCommand(query, con);
                            command.Parameters.AddWithValue("@p1", textBox1.Text);
                            command.Parameters.AddWithValue("@p2", textBox2.Text);
                            command.Parameters.AddWithValue("@p3", textBox3.Text);
                            command.Parameters.AddWithValue("@p4", textBox4.Text);
                            command.Parameters.AddWithValue("@p5", textBox5.Text);
                            command.Parameters.AddWithValue("@id", ID);
                            command.ExecuteNonQuery();
                            con.Close();
                            loadDGView(Queris.selectMeneger);
                        }
                        break;
                }

        }

        public void Delete()
        {
            if (MessageBox.Show("Уверенны что хотите удалить эту запись?", "Удалить?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString());
                switch (Text)
                {
                    case "Договор":
                        {
                            con.Open();
                            OleDbCommand cmd = con.CreateCommand();
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = "DELETE FROM Договор WHERE КодДоговора=@id";
                            cmd.Parameters.AddWithValue("@id", ID);
                            cmd.ExecuteNonQuery();
                            con.Close();
                            loadDGView(Queris.selectDogovor);
                        }
                        break;
                    case "Переговоры":
                        {
                            con.Open();
                            OleDbCommand cmd = con.CreateCommand();
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = "DELETE FROM Переговоры WHERE КодПереговоров=@id";
                            cmd.Parameters.AddWithValue("@id", ID);
                            cmd.ExecuteNonQuery();
                            con.Close();
                            loadDGView(Queris.selectPeregovory);
                        }
                        break;
                    case "Клиент":
                        {
                            con.Open();
                            OleDbCommand cmd = con.CreateCommand();
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = "DELETE FROM Клиент WHERE КодКлиента=@id";
                            cmd.Parameters.AddWithValue("@id", ID);
                            cmd.ExecuteNonQuery();
                            con.Close();
                            loadDGView(Queris.selectKlient);
                        }
                        break;
                    case "Услуги":
                        {
                            con.Open();
                            OleDbCommand cmd = con.CreateCommand();
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = "DELETE FROM Услуги WHERE КодУслуги=@id";
                            cmd.Parameters.AddWithValue("@id", ID);
                            cmd.ExecuteNonQuery();
                            con.Close();
                            loadDGView(Queris.selectUslugi);
                        }
                        break;
                    case "Менеджер":
                        {
                            con.Open();
                            OleDbCommand cmd = con.CreateCommand();
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = "DELETE FROM Менеджер WHERE КодМенеджера=@id";
                            cmd.Parameters.AddWithValue("@id", ID);
                            cmd.ExecuteNonQuery();
                            con.Close();
                            loadDGView(Queris.selectMeneger);
                        }
                        break;
                    case "Отчет":
                        {
                            con.Open();
                            OleDbCommand cmd = con.CreateCommand();
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = "DELETE FROM Отчет WHERE КодОтчета=@id";
                            cmd.Parameters.AddWithValue("@id", ID);
                            cmd.ExecuteNonQuery();
                            con.Close();
                            loadDGView(Queris.selectOtchet);
                        }
                        break;
                }
            }
        }
        private void удалениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Delete();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void поискToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SearchBar.Visible = true;
        }

        private void NameText_Enter(object sender, EventArgs e)
        {
            if (this.Text == "Услуги")
            {
                if (SearchBar.Text == "Поиск по имени")
                {
                    SearchBar.Text = "";
                    SearchBar.ForeColor = Color.Black;
                }
            }
            else
                if (this.Text == "Клиент")
            {
                if (SearchBar.Text == "Поиск по имени")
                {
                    SearchBar.Text = "";
                    SearchBar.ForeColor = Color.Black;
                }
            }
            else
                if (this.Text == "Менеджер")
            {
                if (SearchBar.Text == "Поиск по имени")
                {
                    SearchBar.Text = "";
                    SearchBar.ForeColor = Color.Black;
                }
            }

        }

        private void NameText_TextChanged(object sender, EventArgs e)
        {
            if (this.Text == "Услуги")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                if (SearchBar.Text != "")
                {
                    cmd.CommandText = "SELECT * FROM Услуги Where (ВидУслуги like '%" + SearchBar.Text + "%')";
                }
                else
                {
                    if ((SearchBar.Text == "Поиск по имени") || (SearchBar.Text == ""))
                    {
                        cmd.CommandText = Queris.selectUslugi;
                    }
                }
                cmd.ExecuteNonQuery();
                con.Close();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            else
            if (this.Text == "Менеджер")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                if (SearchBar.Text != "")
                {
                    cmd.CommandText = "SELECT * FROM Менеджер Where (Фамилия like '%" + SearchBar.Text + "%')";
                }
                else
                {
                    if ((SearchBar.Text == "Поиск по Менеджеру") || (SearchBar.Text == ""))
                    {
                        cmd.CommandText = Queris.selectMeneger;
                    }
                }
                cmd.ExecuteNonQuery();
                con.Close();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            else
            if (this.Text == "Клиент")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                if (SearchBar.Text != "")
                {
                    cmd.CommandText = "SELECT * FROM Клиент Where (НазваниеКомпании like '%" + SearchBar.Text + "%')";
                }
                else
                {
                    if ((SearchBar.Text == "Поиск по Клиенту") || (SearchBar.Text == ""))
                    {
                        cmd.CommandText = Queris.selectKlient;
                    }
                }
                cmd.ExecuteNonQuery();
                con.Close();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
        }
        private void NameText_Leave(object sender, EventArgs e)
        {
            if (SearchBar.Text == "")
            {
                SearchBar.Text = "Поиск по имени";
                SearchBar.ForeColor = Color.Silver;
            }
        }
        private void файлToolStripMenuItem_Click(object sender, EventArgs e)
        {


        }

        private void выборТаблицToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ChooseTable ch = new ChooseTable();
            ch.Show();
            this.Hide();
        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            chooseTable.Visible = true;
        }

        private void выводВExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = null;
            Excel.Workbook excelWorkBook = null;
            try
            {
                excelApp = new Excel.Application();
                Excel.Worksheet excelWorkSheet;
                excelWorkBook = excelApp.Workbooks.Add(System.Reflection.Missing.Value);
               // excelWorkSheet = (Excel.Worksheet)excelWorkBook.Worksheets.get_Item(1);
                excelApp.Columns.ColumnWidth = 20;

                int colAmount = dataGridView1.Columns.Count;

                for (int column = 0; column < colAmount; column++)
                    excelApp.Cells[1, column + 1] = dataGridView1.Columns[column].HeaderText;

                for (int row = 0; row < dataGridView1.RowCount; row++)
                    for (int column = 0; column < dataGridView1.ColumnCount; column++)
                    {
                        string value = dataGridView1.Rows[row].Cells[column].Value?.ToString() ?? "";
                        excelApp.Cells[row + 2, column + 1] = value; ;
                    }

                excelApp.Visible = true;
                excelApp.UserControl = true;
            }
            catch (Exception)
            {
                MessageBox.Show("Не удалось открыть Excel!");
                excelWorkBook?.Close(0);
                excelApp?.Quit();
            }
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.SelectionChanged -= dataGridView1_SelectionChanged;
            if (selected > -1)
            {
                loadDGView(Queris.topManager);
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0 && dataGridView1.CurrentRow.Cells[0].Value.ToString() != "")
                selected = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString());
        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.SelectionChanged -= dataGridView1_SelectionChanged;
            if (selected > -1)
            {
                loadDGView(Queris.processPeregovory);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //loadDGView(Queris.selectUniver);
            dataGridView1.SelectionChanged += dataGridView1_SelectionChanged;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns["КодГорода"].Visible = false;
        }

        private void печатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var printDialog = new PrintDialog();
            printDialog.ShowDialog();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            loadDGView(Queris.produktDogovory);
        }
    }
    class CbbEntry
    {
        public int ID { get; set; }
        public string Field { get; set; }
        public CbbEntry(int id, string field)
        {
            ID = id;
            Field = field;
        }
    }
}