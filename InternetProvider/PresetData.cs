using System;
using System.Collections.Generic;

namespace InternetProvider
{
    internal class PresetData
    {
        static List<string> _name = new List<string> { "Олександр", "Олександра", "Олег", "Ірина", "Ксенія", "Дмитрий", "Ярослав", "Ольга" };
        static List<string> _surname = new List<string> { "Селезньов", "Аксютін", "Смирнов", "Кузнєцов", "Попов", "Васильєв", "Петров", "Соколів", "Іванов", "Михайлов", "Зайцев", "Карпов" };

        static List<string> street = new List<string> { "П`ята авеню", "Бродвей", "Арбат", "Уолл Стріт", "Вулиця Червоних ліхтарів", "Бульвар машинобудівників", "Дерибасівська вулиця" };

        private static Random random = new Random();

        public static int GetRandomValue(int max, int min = 0)
        {
            return random.Next(min, max);
        }

        public static string GetFullName()
        {
            string name = GetRandomString(_name);
            string surname = GetRandomString(_surname) + ((name[name.Length-1] == 'а' || name[name.Length - 1] == 'я') ? "а" : "");
            return $"{surname} {name}";
        }

        public static string GetStreet()
        {
            return GetRandomString(street);
        }

        private static string GetRandomString(List<string> text)
        {
            return text[random.Next(0, text.Count)];
        }

        public static string GetPhoneNumber()
        {
            return $"+380{random.Next(100000000, 999999999)}";
        }

    }
}
