using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;

namespace Reg
{
    public partial class RegistrationForm : Form
    {
        /* Переменные, которые будут хранить на протяжение работы программы логин и пароль. */
        public string login = string.Empty;
        public string password = string.Empty;
        private Users user = new Users(); // Экземпляр класса пользователей.

        public RegistrationForm()
        {
            InitializeComponent();

            LoadUsers(); // Метод десериализует класс.
        }

        private void LoadUsers()
        {
           
            try
            {
                FileStream fs = new FileStream("Users.dat", FileMode.Open);

                BinaryFormatter formatter = new BinaryFormatter();

                user = (Users)formatter.Deserialize(fs);

                fs.Close();
            }
            catch { return; }
        }

        private void EnterToForm()
        {
            //for (int i = 0; i < user.Logins.Count; i++) // Ищем пользователя и проверяем правильность пароля.
            //{
            //    if (user.Logins[i] == loginTextBox.Text && user.Passwords[i] == passwordTextBox.Text)
            //    {
            //        login = user.Logins[i];
            //        password = user.Passwords[i];

            //        //MessageBox.Show("Вы вошли в систему!");

            //        this.Close();
            //    }
            //    else if (user.Logins[i] == loginTextBox.Text && passwordTextBox.Text != user.Passwords[i])
            //    {
            //        login = user.Logins[i];

            //        MessageBox.Show("Неверный пароль!");
            //    }
            //}

            //if (login == "") { MessageBox.Show("Пользователь " + loginTextBox.Text + " не найден!"); }

            if (loginTextBox.Text == "admin" && passwordTextBox.Text == "admin")
                Close();

        }

        private void AddUser() // Регистрируем нового пользователя.
        {
            if (loginTextBox.Text == "" || passwordTextBox.Text == "") { MessageBox.Show("Не введен логин или пароль!"); return; }

            user.Logins.Add(loginTextBox.Text);
            user.Passwords.Add(passwordTextBox.Text);

            FileStream fs = new FileStream("Users.dat", FileMode.OpenOrCreate);

            BinaryFormatter formatter = new BinaryFormatter();
            formatter.Serialize(fs, user); // Сериализуем класс.

            fs.Close();

            login = loginTextBox.Text;

            this.Close();
        }

        private void exitButton_Click(object sender, EventArgs e)
        {
            Application.Exit(); // Закрываем программу.
        }

        private void regButton_Click(object sender, EventArgs e)
        {
            AddUser();
            
        }

        private void enterButton_Click(object sender, EventArgs e)
        {
            EnterToForm();
        }

        private void RegistrationForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (login == "" | password == "") { Application.Exit(); }
        }

        private void RegistrationForm_Load(object sender, EventArgs e)
        {

        }
    }
}
