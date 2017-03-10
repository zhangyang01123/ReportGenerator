using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Data;
using System.Reflection;
using Excel.Object;


namespace LocalIMEUsage
{
    public class ExcelWrite : ExcelFileWriter<ExcelObjectElement>
    {
        public object[,] myExcelData;
        private int myRowCnt;
        public override object[] Headers
        {
            get
            {
                return columnHeaders;
            }
        }

        public override void FillRowData(List<ExcelObjectElement> list)
        {

            myRowCnt = list.Count;
            myExcelData = new object[RowCount + 1, ColumnCount];
            for (int row = 1; row <= myRowCnt; row++)
            {
                myExcelData[row, 0] = list[row - 1].CodePage;
                myExcelData[row, 1] = list[row - 1].Language;
                myExcelData[row, 2] = list[row - 1].LocalIMEOnCount;
                myExcelData[row, 3] = list[row - 1].LocalIMEOffCount;
                myExcelData[row, 4] = list[row - 1].Percentage;
            }
        }

        public override object[,] ExcelData
        {
            get
            {
                return myExcelData;
            }
        }

        public override int ColumnCount
        {
            get
            {
                return columnCount;
            }
        }

        public override int RowCount
        {
            get
            {
                return myRowCnt;
            }
        }
    }

    class LocalIMEVSUILanguageObject
    {
        public string localIME;
        public string language;
        public int count;
    }

    public class ExcelObjectElement
    {
        public string CodePage;
        public string Language;
        public int LocalIMEOnCount;
        public int LocalIMEOffCount;
        public double Percentage;

        public ExcelObjectElement(string codepage, string language, int onCount, int offCount, double percentage)
        {
            CodePage = codepage;
            Language = language;
            LocalIMEOnCount = onCount;
            LocalIMEOffCount = offCount;
            Percentage = percentage;
        }
    }
    
    class Program
    {
        static void Main(string[] args)
        {
            Dictionary<string, int> onLocalIMEVSUILanguageDictionary = new Dictionary<string, int>();
            Dictionary<string, int> offLocalIMEVSUILanguageDictionary = new Dictionary<string, int>();
            List<LocalIMEVSUILanguageObject> localIMEVSUILanguageObjectList = new List<LocalIMEVSUILanguageObject>();
            List<ExcelObjectElement> excelObjectList = new List<ExcelObjectElement>();
            StreamReader sr = new StreamReader(@"LocalIMEAndUILanguage.txt");
            Dictionary<string, string> codepageToLanguage = new Dictionary<string, string>();

            while (sr.ReadLine() != null)
            {
                string line = null;
                LocalIMEVSUILanguageObject newObject = new LocalIMEVSUILanguageObject();
                line = sr.ReadLine();
                string[] strArray = line.Split(',');
                newObject.localIME = strArray[0].Trim() ;
                line = sr.ReadLine();
                strArray = line.Split(',');
                
                if(strArray[0].Trim() != "null")
                {
                    int languageID = Int32.Parse(strArray[0].Trim());
                    string language = GetLocaleName(languageID);
                    if (codepageToLanguage.ContainsKey(strArray[0].Trim()) == false)
                    {
                        codepageToLanguage.Add(strArray[0].Trim(), language);
                    }
                }
                newObject.language = strArray[0].Trim();

                string test = sr.ReadLine();
                if (Int32.TryParse(test, out newObject.count) == false)
                {
                    newObject.count = 0;
                }
                localIMEVSUILanguageObjectList.Add(newObject);
                if(newObject.localIME == "0")
                {
                    offLocalIMEVSUILanguageDictionary.Add(newObject.language, newObject.count);
                }
                else if (newObject.localIME == "1")
                {
                    onLocalIMEVSUILanguageDictionary.Add(newObject.language, newObject.count);
                }
                sr.ReadLine();
            }
            sr.Close();

            foreach (var item in offLocalIMEVSUILanguageDictionary)
            {
                if(onLocalIMEVSUILanguageDictionary.ContainsKey(item.Key) == false)
                {
                    onLocalIMEVSUILanguageDictionary.Add(item.Key, 0);
                }
            }

            foreach (var item in offLocalIMEVSUILanguageDictionary)
            {
                if (item.Key == "null")
                    continue;
                if (onLocalIMEVSUILanguageDictionary.ContainsKey(item.Key))
                {
                    Console.WriteLine("{0}  OFF:{1}  ON:{2}", item.Key, item.Value, onLocalIMEVSUILanguageDictionary[item.Key]);
                }
                else
                {
                    Console.WriteLine("{0}  OFF:{1}  ON:0", item.Key, item.Value);
                }
            }

            List<KeyValuePair<string, int>> lst = new List<KeyValuePair<string, int>>(onLocalIMEVSUILanguageDictionary);
            lst.Sort(delegate (KeyValuePair<string, int> s1, KeyValuePair<string, int> s2)
            {
                return s2.Value.CompareTo(s1.Value);
            });

            foreach (KeyValuePair<string, int> kvp in lst)
                Console.WriteLine(kvp.Key + "：" + kvp.Value);

            foreach(var item in lst)
            {
                if (item.Key == "null")
                    continue;
                if (offLocalIMEVSUILanguageDictionary.ContainsKey(item.Key))
                {
                    Console.WriteLine("{0}  ON:{1} OFF:{2}, {3}%", item.Key, item.Value, offLocalIMEVSUILanguageDictionary[item.Key], 1.0*item.Value / offLocalIMEVSUILanguageDictionary[item.Key] * 100);
                }
                else
                {
                    Console.WriteLine("{0}  ON:{1}  OFF:0", item.Key, item.Value);
                }
            }

            FileStream fs = new FileStream(@"LocalIMECEIPReport.txt", FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);

            foreach (var item in lst)
            {
                if (item.Key == "null")
                    continue;
                if (offLocalIMEVSUILanguageDictionary.ContainsKey(item.Key))
                {
                    Console.WriteLine("{0}  ON:{1} OFF:{2}, {3}%", item.Key, item.Value, offLocalIMEVSUILanguageDictionary[item.Key], 1.0 * item.Value / offLocalIMEVSUILanguageDictionary[item.Key] * 100);
                    sw.WriteLine("{0}\t{1}\t{2}", item.Key, item.Value, offLocalIMEVSUILanguageDictionary[item.Key]);
                    excelObjectList.Add(new ExcelObjectElement(item.Key, codepageToLanguage[item.Key], item.Value, offLocalIMEVSUILanguageDictionary[item.Key], 1.0 * item.Value / offLocalIMEVSUILanguageDictionary[item.Key] * 10000));
                }
                else
                {
                    Console.WriteLine("{0}  ON:{1}  OFF:0", item.Key, item.Value);
                    sw.WriteLine("{0}\t{1}\t0", item.Key, item.Value);
                    excelObjectList.Add(new ExcelObjectElement(item.Key, codepageToLanguage[item.Key], item.Value, 0, 0));
                }
            }

            sw.Close();
            fs.Close();

            object[] headerNames = { "CodePage", "Language", "Local IME ON", "Local IME OFF", "On/Off (*10000)" };
            ExcelFileWriter<ExcelObjectElement> myExcel = new ExcelWrite();
            myExcel.WriteDateToExcel(@"C:\TEMP\LocalIMEUsageStatistics.xls", excelObjectList, headerNames.Count(), headerNames, "A1", "E1");
        }

        static string GetLocaleName(int cid)
        {
            var buffer = new StringBuilder(256);
            if (GetLocaleInfo(cid, 6, buffer, buffer.Capacity) == 0)
            {
                throw new System.ComponentModel.Win32Exception();
            }

            return buffer.ToString();
        }

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int GetLocaleInfo(int LCID, int LCType, StringBuilder buffer, int buflen);
    }
}
