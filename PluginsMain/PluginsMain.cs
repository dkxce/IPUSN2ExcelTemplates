//
// Основной файл по общей информации для работы плагинов
//

using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Serialization;


namespace PluginsMain
{
    #region XML FILES

    [XmlRoot("PluginInfo")]
    public class PluginInfo : XMLSaved<PluginInfo>
    {
        public string PluginType;
        public string PluginName;
        public string SupportedDocuments;
        public string PluginShortName;
        public string Executable;
        public string Library;

        [XmlIgnore]
        public List<string> SupportedDocumentsAsList
        {
            get
            {
                return new List<string>(SupportedDocuments.Split(','));
            }
        }

        [XmlIgnore]
        public bool RunAsExe
        {
            get { return !string.IsNullOrEmpty(Executable); }
        }

        [XmlIgnore]
        public bool RunAsDLL
        {
            get { return !string.IsNullOrEmpty(Library); }
        }

        [XmlIgnore]
        public string XML
        {
            get
            {
                return PluginInfo.Save(this);
            }
        }

        public static List<PluginInfo> ScanForPlugins(string path = null)
        {
            if (string.IsNullOrEmpty(path)) path = XMLSaved<PluginInfo>.CurrentDirectory();
            List<PluginInfo> result = new List<PluginInfo>();
            foreach(string f in Directory.GetFiles(path, "_plugin_config.xml", SearchOption.AllDirectories))
            {
                PluginInfo info = PluginInfo.Load(f);
                if (!string.IsNullOrEmpty(info.Executable)) info.Executable = Path.Combine(Path.GetDirectoryName(f), info.Executable);
                if (!string.IsNullOrEmpty(info.Library)) info.Library = Path.Combine(Path.GetDirectoryName(f), info.Library);
                result.Add(info);
            };
            return result;
        }
    }

    [XmlRoot("Body")]
    public class ExportedDocument: XMLSaved<ExportedDocument>
    {
        [XmlElement("Type")]
        public string DocType = "счет";

        [XmlArray("Document"), XmlArrayItem("Field")]
        public List<Field> DocFields = new List<Field>();

        [XmlArray("Items"), XmlArrayItem("Item")]
        public List<DocumentItem> DocItems = new List<DocumentItem>();

        public string GetDocField(string id)
        {
            id = id.Trim('%', '$', ' ');
            if (DocFields.Count == 0) return "";
            foreach(Field f in DocFields)
            {
                string fid = f.id.Trim('%', '$', ' ');
                if (fid == id) return f.value;
            };
            return "";
        }
    }

    public class DocumentItem
    {
        [XmlElement("Field")]
        public List<Field> DocFields = new List<Field>();
    }

    public class Field
    {
        [XmlAttribute]
        public string id;
        [XmlText]
        public string value;
    }

    #endregion XML FILES

    #region PLUGIN INTERFACE

    public interface IPluginInterface
    {
        void SetApplicationDirectory(string path);
        void SetPluginDirectory(string path);
        void OpenConfig();

        void Run(string filePath, string args);        

        string GetPluginType();
        string GetPluginName();
        string GetSupportedDocuments();
        string GetPluginShortName();
    }

    public class PluginLoader
    {
        public static IPluginInterface LoadDLL(string path)
        {
            System.Reflection.Assembly asm = System.Reflection.Assembly.LoadFile(path);
            Type[] tps = asm.GetTypes();
            Type asmType = null;
            foreach (Type tp in tps) if (tp.GetInterface(typeof(IPluginInterface).ToString()) != null) asmType = tp;

            System.Reflection.ConstructorInfo ci = asmType.GetConstructor(new Type[] { });
            try
            {
                return (IPluginInterface)ci.Invoke(new object[] { });
            }
            catch (Exception ex)
            {
                string rr = GetReference(path);
                if (rr != "") rr = "\r\nAssembly references: " + rr;
                throw new Exception(" Couldn't load assembly " + System.IO.Path.GetFileName(path) + " - " + ex.Message + rr);
            };
        }

        private static string GetReference(string path)
        {
            try
            {
                System.Reflection.Assembly asm = System.Reflection.Assembly.LoadFile(path);
                Type[] tps = asm.GetTypes();
                Type asmType = null;
                foreach (Type tp in tps)
                    if (tp.Name == "References") asmType = tp;
                System.Reflection.MethodInfo mi = asmType.GetMethod("Reference");
                return (string)mi.Invoke(null, null);
            }
            catch
            {
                return "";
            };
        }
    }

    #endregion PLUGIN INTERFACE
}
