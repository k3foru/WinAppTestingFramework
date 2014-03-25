//===============================================================================
// Ellis WinApp Testing Framework Library
// By Kiran Kumar
//===============================================================================

using System;

namespace Ellis.WinApp.Testing.Framework
{
    public class Generator : AppContext
    {
        public static string GenerateNewName(string name)
        {
            var n = new Random().Next(9999);
            var newName = name + n;
            return newName;
        }

        public static string GenerateSSNNumber()
        {
            var rnd = new Random();
            var num = rnd.Next(100000000, 999999999).ToString();
            return num;
        }

        public static string GenerateRandomNumber(int length)
        {
            var rnd = new Random();
            var num = rnd.Next(0, length).ToString();
            return num;
        }
    }
}