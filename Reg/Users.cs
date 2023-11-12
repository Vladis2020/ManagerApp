using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.Serialization.Formatters.Binary;

namespace Reg
{
    [Serializable]
    public class Users // Класс, который хранит регистрационные данные.
    {
        public List<string> Logins = new List<string>(); // Логин.
        public List<string> Passwords = new List<string>(); // Пароль.
    }
}
