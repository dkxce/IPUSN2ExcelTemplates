//
// Запуск EXE плагина
//

using System;
using System.IO;
using System.Collections.Generic;
using System.Xml;
using PluginsMain;

namespace ExcelTemplates
{
    internal static class Program
    {
        private static List<string> args = new List<string>();

        [STAThread]
        static void Main(string[] a)
        {
            args = new List<string>(a);

            if (args.Contains("/create_plugins_config")) CreatePluginsConfig();
            if (args.Contains("/list_plugins")) ListPlugins();

            if (args.Contains("/create_test_bill_doc")) CreateTestDoc(@"Tests\test_doc_bill.xml", "счет");
            if (args.Contains("/create_test_act_doc")) CreateTestDoc(@"Tests\test_doc_act.xml", "акт");

            if (args.Contains("/install")) Install();
            if (args.Contains("/config")) OpenConfig();
            if (args.Count > 0) ProcessFiles(); else Help();
        }

        private static void ProcessFiles()
        {
            foreach (string f in args)
            {
                if (f.StartsWith("/")) continue;
                try { if (!File.Exists(f)) continue; } catch { continue; };
                ProcessFile(f);
            };
        }

        private static void CreatePluginsConfig()
        {
            PluginInfo pluginsMain = new PluginInfo()
            {
                PluginType = "export_documents",
                PluginName = "Вывод документов на основе расширенных шаблонов Excel",
                SupportedDocuments = "счет,акт",
                PluginShortName = "Вывод Excel Templates",
                Executable = "ExcelTemplates.exe",
                Library = "ExcelTemplatesLib.dll"
            };
            PluginInfo.SaveHere("_plugin_config.xml", pluginsMain);
        }

        private static void ListPlugins()
        {
            WinConsoleApplication.Initialize(false, true, false);

            foreach (PluginInfo info in PluginInfo.ScanForPlugins())
                Console.WriteLine(info.XML);

            System.Threading.Thread.Sleep(500); // Requires for stdout
        }

        private static void Help()
        {
            WinConsoleApplication.Initialize(false, true, false);
            
            Console.WriteLine((new ExcelTemplatesLib.ExcelTemplatesLib()).GetPluginName());
            Console.WriteLine();
            Console.WriteLine("Usage: ");
            Console.WriteLine(" /install");
            Console.WriteLine(" /config");
            Console.WriteLine(" <document_file>");
            Console.WriteLine();
            Console.WriteLine("Test Only: ");
            Console.WriteLine(" /list_plugins");
            Console.WriteLine(" /create_test_bill_doc");
            Console.WriteLine(" /create_test_act_doc"); 
            
            System.Threading.Thread.Sleep(5000); // Requires for stdout
        }

        private static void Install()
        {
            WinConsoleApplication.Initialize(false, true, false);
            Installer.Install();
            System.Threading.Thread.Sleep(500); // Requires for stdout
        }
        
        private static void OpenConfig()
        {
            IPluginInterface pi = PluginLoader.LoadDLL(Path.Combine(XMLSaved<int>.CurrentDirectory(), "ExcelTemplatesLib.dll"));
            pi.OpenConfig();
        }

        private static void ProcessFile(string fileName)
        {
            IPluginInterface pi = new ExcelTemplatesLib.ExcelTemplatesLib();
            pi.Run(fileName, null);
        }

        private static void CreateTestDoc(string fileName, string docType = "счет")
        {
            ExportedDocument xd = new ExportedDocument()
            {
                DocType = docType,
                DocFields = new List<Field>()
                {
                    new Field(){ id = "MYBANK", value = "TEST_MYBANK" },
                    new Field(){ id = "MYINN", value = "TEST_MYINN" },
                    new Field(){ id = "MYCOMPANY", value = "TEST_MYCOMPANY" },
                    new Field(){ id = "MYBIK", value = "TEST_MYBIK" },
                    new Field(){ id = "MYKS", value = "TEST_MYKS" },
                    new Field(){ id = "MYACCOUNT", value = "TEST_MYACCOUNT" },
                    new Field(){ id = "NUMBER", value = "TEST_NUMBER" },
                    new Field(){ id = "DATE", value = "TEST_DATE" },
                    new Field(){ id = "BYCONTRACT", value = "TEST_BYCONTRACT" },
                    new Field(){ id = "MYINN", value = "TEST_MYINN" },
                    new Field(){ id = "MYCOMPANY", value = "TEST_MYCOMPANY" },
                    new Field(){ id = "MYADDRESS", value = "TEST_MYADDRESS" },
                    new Field(){ id = "PARTNER", value = "TEST_PARTNER" },
                    new Field(){ id = "PARTNER_NAME", value = "TEST_PARTNER_NAME" },
                    new Field(){ id = "PARTNER_INN", value = "" },
                    new Field(){ id = "PARTNER_KPP", value = "" },
                    new Field(){ id = "PARTNER_ADDRESS", value = "TEST_PARTNER_ADDRESS" },
                    new Field(){ id = "RESULT", value = "12 345,67" },
                    new Field(){ id = "NN", value = "TEST_NN" },
                    new Field(){ id = "CURR", value = "TEST_CURR" },
                    new Field(){ id = "RESULT_DESC", value = "TEST_RESULT_DESC" },
                    new Field(){ id = "MYSHORTNAME", value = "TEST_MYSHORTNAME" }
                },
                DocItems = new List<DocumentItem>()
                {
                    new DocumentItem()
                    {
                        DocFields = new List<Field>()
                        {
                            new Field(){ id = "N", value = "TEST_N1" },
                            new Field(){ id = "CODE_AND_NAME", value = "TEST_CODE_AND_NAME1" },
                            new Field(){ id = "TITLE", value = "TEST_TITLE1" },
                            new Field(){ id = "CNT", value = "TEST_CNT1" },
                            new Field(){ id = "U", value = "TEST_U1" },
                            new Field(){ id = "PRICE", value = "TEST_PRICE1" },
                            new Field(){ id = "SUM", value = "TEST_SUM1" }
                        }
                    },
                    new DocumentItem()
                    {
                        DocFields = new List<Field>()
                        {
                            new Field(){ id = "N", value = "TEST_N2" },
                            new Field(){ id = "CODE_AND_NAME", value = "TEST_CODE_AND_NAME2" },
                            new Field(){ id = "TITLE", value = "TEST_TITLE2" },
                            new Field(){ id = "CNT", value = "TEST_CNT2" },
                            new Field(){ id = "U", value = "TEST_U2" },
                            new Field(){ id = "PRICE", value = "TEST_PRICE2" },
                            new Field(){ id = "SUM", value = "TEST_SUM2" }
                        }
                    }
                }
            };
            ExportedDocument.SaveHere(fileName, xd);
        }

    }
}
