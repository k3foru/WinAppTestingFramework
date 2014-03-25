//===============================================================================
// Ellis WinApp Testing Framework Library
// By Kiran Kumar
//===============================================================================

using System.IO;

namespace Ellis.WinApp.Testing.Framework
{
    public class CSVReader
    {
        public static string ControlType;
        public static string LocatorType;
        public static string LocatorValue;

        public static void ReadDataFromCsv(string fileName, string element)
        {
            var streamReader = new StreamReader(File.OpenRead(fileName));
            while (!streamReader.EndOfStream)
            {
                var str = streamReader.ReadLine();

                if (str == null) continue;
                var strArray1 = str.Split(new[] {','});

                if (!strArray1[0].Contains(element)) continue;
                var strArray2 = strArray1[1].Split(new[] {'|'});

                ControlType = strArray2[0];
                LocatorType = strArray2[1];
                LocatorValue = strArray2[2];
            }
        }
    }
}
