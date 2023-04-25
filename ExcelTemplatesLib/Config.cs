using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace ExcelTemplatesLib
{
    [Serializable]
    public class PluginConfig: XMLSaved<PluginConfig>
    {
        public byte StartMode = 0;
        public bool QRIP = false;
        public bool MatrixBar = false;
        public bool Code39Bar = false;
        public string MatrixCode = "Data Matrix";
        public string SingleCode = "Code 39";
        [XmlArrayItem("Document")]
        public List<DocumentFile> LastTemplates = new List<DocumentFile>();

        public static string FileName { get { return Path.Combine(XMLSaved<int>.CurrentDirectory(), "config.xml"); } }

        public class DocumentFile
        {
            [XmlAttribute("doc")]
            public string d;
            [XmlAttribute("file")]
            public string f;

            public DocumentFile() { }
            public DocumentFile(string d, string f) { this.d = d; this.f = f; }
        }
    }
}
