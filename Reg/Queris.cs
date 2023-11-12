﻿using System;
using System.Collections.Generic;
using System.Text;

namespace Reg
{
    class Queris
    {
        public static string selectUslugi = "Select * from Услуги";
        public static string selectDogovor = "SELECT Договор.КодДоговора, Договор.КодКлиента, Клиент.НазваниеКомпании AS Клиент, Договор.КодУслуги, Услуги.ВидУслуги AS Услуга, Договор.КодМенеджера, Менеджер.Фамилия AS Менеджер, Договор.Дата FROM (((Договор INNER JOIN Клиент ON Договор.КодКлиента = Клиент.КодКлиента) INNER JOIN Услуги ON Договор.КодУслуги = Услуги.КодУслуги) INNER JOIN Менеджер ON Договор.КодМенеджера = Менеджер.КодМенеджера)";
        public static string selectKlient= "select Клиент.КодКлиента, Клиент.НазваниеКомпании, Клиент.КодУслуги, Услуги.ВидУслуги as Услуги from Клиент inner join Услуги on Клиент.КодУслуги = Услуги.КодУслуги";
        public static string selectPeregovory= "SELECT Переговоры.КодПереговоров, Переговоры.КодУслуги, Переговоры.КодКлиента, Переговоры.КодМенеджера,Переговоры.КонтактныйАдрес, Переговоры.Статус, Услуги.ВидУслуги as Услуги, Клиент.НазваниеКомпании AS Клиент, Менеджер.Фамилия AS Менеджер FROM ((Переговоры INNER JOIN Услуги ON Переговоры.КодУслуги = Услуги.КодУслуги) INNER JOIN Клиент ON Переговоры.КодКлиента = Клиент.КодКлиента) INNER JOIN Менеджер ON Переговоры.КодМенеджера = Менеджер.КодМенеджера";
        public static string selectMeneger= "select * from Менеджер";
        public static string selectOtchet = "select Отчет.КодОтчета, Отчет.КодМенеджера, Менеджер.Фамилия as Менеджер, Отчет.КоличествоЗаключенныхДоговоров from Отчет inner join Менеджер on Отчет.КодОтчета = Менеджер.КодМенеджера";
        public static string topManager = "Select Top 1 * From Отчет INNER JOIN Менеджер ON Отчет.КодМенеджера = Менеджер.КодМенеджера ORDER BY Отчет.КоличествоЗаключенныхДоговоров DESC";
        public static string processPeregovory = "SELECT Переговоры.КодПереговоров, Переговоры.КодУслуги, Переговоры.КодКлиента, Переговоры.КодМенеджера,Переговоры.КонтактныйАдрес, Переговоры.Статус, Услуги.ВидУслуги as Услуги, Клиент.НазваниеКомпании AS Клиент, Менеджер.Фамилия AS Менеджер FROM ((Переговоры INNER JOIN Услуги ON Переговоры.КодУслуги = Услуги.КодУслуги) INNER JOIN Клиент ON Переговоры.КодКлиента = Клиент.КодКлиента) INNER JOIN Менеджер ON Переговоры.КодМенеджера = Менеджер.КодМенеджера Where Статус = 'В процессе'";
        public static string produktDogovory = "Select * From Договор Inner Join Услуги ON Договор.КодУслуги = Услуги.КодУслуги Where Услуги.КодУслуги = 1";
    }
}
