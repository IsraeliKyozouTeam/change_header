using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Xml;
using System.Configuration;
using System.Xml.Linq;



namespace change_header
{
    class Program
    {
        public static Dictionary<string, string> dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public static void InitDictionary()
        {
             XmlReader reader = XmlReader.Create(ConfigurationSettings.AppSettings["headersmap"]);
            while(reader.Read())
            {
                if (reader.Name=="header")
                {
                    while (reader.MoveToNextAttribute())
                    {// Read the attributes.
                        string x = reader.Value;
                        dict.Add(x,"");
                        reader.MoveToNextAttribute();
                        dict[x] = reader.Value;
                        
                    }
                }
            }
        }
        
            static void Main(string[] args)
        {
            InitDictionary();
            string url = ConfigurationSettings.AppSettings["filepath"];
            if (url.Contains("xlsx") || url.Contains("xls"))
            {
                Console.WriteLine("please enter-two tabs for two tabs");
                Console.WriteLine("if you want just one tab, please enter something else");
                if (Console.ReadLine() == "two tabs")
                {
                    exel x = new exel(url);
                    x.openAnotherFile();
                    x.changeDataOfFile(dict);
                    x.closeFile();
                    Console.WriteLine(x.xlApp);
                    Console.ReadKey();
                }
                else
                {
                    exel x = new exel(url);
                    x.changeDataOfFile(dict);
                    x.closeFile();
                    Console.WriteLine(x.xlApp);
                    Console.ReadKey();
                }
            }
            if (url.Contains("csv")||url.Contains("txt"))
            {
                exel x = new exel(url);
                x.changeDataOfFile(dict);
                x.closeFile();
                Console.WriteLine(x.xlApp);
                Console.ReadKey();
            }
        }



    }
}
