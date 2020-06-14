using System;
using System.Collections.Generic;
using System.Linq;
using COMAdmin;
using System.IO;
using Microsoft.Win32;


namespace COM1CMR
{
    internal class Program
    {
        class Info1C
        {
            public string ComPath { get; set; }
            public string Ver { get; set; }
            public string Clsid { get; set; }

            public Info1C (string comPathc, string verc)
            {
                ComPath = comPathc;
                Ver = verc;
            }
        }

        static void Main(string[] args)
        {
            Console.WriteLine("Запущено формирование com оберток.\n");

            try
            {
                CreateNewCom();
                Console.WriteLine("\nФормирование оберток прошло успешно.");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

            Console.WriteLine("Для продолжения нажмите любую клавишу....");
            Console.ReadKey();
        }

        static void CreateNewCom()
        {
            ICatalogObject catalogObject1Cv8 = null;
            ICatalogObject mainConnector = null;

            String dir1C = @"C:\Program Files (x86)\1cv8";

            Console.WriteLine("Каталог для поиска библиотек com: " + dir1C + "\n");

            COMAdminCatalog catalog = new COMAdminCatalogClass();

            ICatalogCollection appCollection = (ICatalogCollection)catalog.GetCollection("Applications");
            appCollection.Populate();

            foreach (ICatalogObject catalogObject in appCollection)
            {
                if (catalogObject.Name.ToString().ToLower() == "1cv8")
                {
                    catalogObject1Cv8 = catalogObject;
                    Console.WriteLine("Каталог компонент 1cv8 найден.\n");
                    break;
                }
            }

            if (catalogObject1Cv8 == null)
            {
                Console.WriteLine("Каталог компонент 1cv8 НЕ найден.\n");
                ICatalogObject new1Cv8App = (ICatalogObject)appCollection.Add();
                new1Cv8App.Value["Name"] = "1cv8";
                new1Cv8App.Value["Activation"] = COMAdminActivationOptions.COMAdminActivationInproc;
                appCollection.SaveChanges();
                Console.WriteLine("Каталог компонент 1cv8 создан\n");

                catalogObject1Cv8 = new1Cv8App;
            }

            if (!Directory.Exists(dir1C))
            {
                Console.WriteLine("Каталог: " + dir1C + " НЕ найден.\n");
                return;
            }

            Console.WriteLine("Каталог: " + dir1C + " найден.\n");

            List <Info1C> verList = new List<Info1C>();

            String[] catalog1C = Directory.GetDirectories(@"C:\Program Files (x86)\1cv8", "8.*");

            foreach (string path in catalog1C)
            {
                String[] arrayName = path.Split('\\');
                var ver = arrayName[arrayName.Length - 1];
                var comPath = path + @"\bin\comcntr.dll";

                if (!File.Exists(comPath))
                {
                    continue;
                }

                Console.WriteLine("Найдена компонента версии: " + ver);  

                verList.Add(new Info1C(comPath, ver));
            }

            Console.WriteLine("\n");

            ICatalogCollection comCollections = (ICatalogCollection)appCollection.GetCollection("Components", catalogObject1Cv8.Key);
            comCollections.Populate();

            foreach (ICatalogObject catalogObject in comCollections)
            {
                if (catalogObject.Name.ToString() == "V83.COMConnector.1")
                {
                    Console.WriteLine("Основная обертка: V83.COMConnector найдена.\n");
                    mainConnector = catalogObject;
                }

                var arrayName = catalogObject.Name.ToString().Split('_');
                var ver = arrayName[arrayName.Length - 1];

                var findVer = verList.Where(s => s.Ver == ver);

                if (findVer.Count() != 0)
                {
                    Console.WriteLine("Найдена обертка для версии: " + findVer.First().Ver);
                    verList.Remove(findVer.First());
                }

            }

            if (mainConnector == null)
            {
                Console.WriteLine("Основная обертка: V83.COMConnector НЕ найдена.");
                catalog.InstallComponent("1cv8", verList.Last().ComPath, "", "");
                Console.WriteLine("Основная обертка: V83.COMConnector создана");
                mainConnector = findMainCom(comCollections);
            }

            foreach (Info1C itemVer in verList)
            {
                Console.WriteLine("\n");
                Console.WriteLine("Создание обертки для версии: " + itemVer.Ver);
                catalog.AliasComponent("1cv8", mainConnector.get_Value("CLSID").ToString(), "", "V83.COMConnector_" + itemVer.Ver, "");
                Console.WriteLine("Обертка для версии: " + itemVer.Ver + " создана.");
            }

            comCollections.Populate();

            Console.WriteLine("\n");
            Console.WriteLine("Начато изменение данных в реестре.");

            foreach (ICatalogObject catalogObject in comCollections)
            {
                if (catalogObject.Name.ToString() == "V83.COMConnector.1")
                {
                    continue;
                }

                var arrayName = catalogObject.Name.ToString().Split('_');
                Info1C findingItem = null;

                try
                {
                    findingItem = verList.First(s => s.Ver == arrayName[arrayName.Length - 1]);
                }
                catch { continue; }

                if (findingItem == null)
                {
                    continue;
                }

                string clsid = catalogObject.get_Value("CLSID").ToString();

                Console.WriteLine("\n");
                Console.WriteLine("Начато изменение данных в реестре, для версии: " + findingItem.Ver);

                RegistryKey readKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Classes\\Wow6432Node\\CLSID\\" + clsid + "\\InprocServer32", true);
                readKey.SetValue("", findingItem.ComPath);
                readKey.Close();

                Console.WriteLine("Произведено изменение данных в реестре, для версии: " + findingItem.Ver);

            }

            comCollections.SaveChanges();
        }

        static ICatalogObject findMainCom(ICatalogCollection comCollections)
        {
            comCollections.Populate();

            foreach (ICatalogObject catalogObject in comCollections)
            {
                if (catalogObject.Name.ToString() == "V83.COMConnector.1")
                {
                    return catalogObject;
                }
            }

            return null;
        }
    }
}
