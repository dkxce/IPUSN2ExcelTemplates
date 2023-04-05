using System;
using System.IO;
using System.Xml;

namespace ExcelTemplatesLib
{
    [Serializable]
    public class PluginConfig: XMLSaved<PluginConfig>
    {
        public byte StartMode = 0;

        public static string FileName { get { return Path.Combine(XMLSaved<int>.CurrentDirectory(), "config.xml"); } }
    }
}
