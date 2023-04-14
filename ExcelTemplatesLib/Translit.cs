using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Text;
using System.Web;

namespace TRANSLIT
{   
    public static class Translit
    {
        private static Dictionary<string, string> words;

        static Translit() { InitDict(); }        

        private static void InitDict()
        {
            words = new Dictionary<string, string>();
            words.Add("а", "a");
            words.Add("б", "b");
            words.Add("в", "v");
            words.Add("г", "g");
            words.Add("д", "d");
            words.Add("е", "e");
            words.Add("ё", "yo");
            words.Add("ж", "zh");
            words.Add("з", "z");
            words.Add("и", "i");
            words.Add("й", "j");
            words.Add("к", "k");
            words.Add("л", "l");
            words.Add("м", "m");
            words.Add("н", "n");
            words.Add("о", "o");
            words.Add("п", "p");
            words.Add("р", "r");
            words.Add("с", "s");
            words.Add("т", "t");
            words.Add("у", "u");
            words.Add("ф", "f");
            words.Add("х", "h");
            words.Add("ц", "c");
            words.Add("ч", "ch");
            words.Add("ш", "sh");
            words.Add("щ", "sch");
            words.Add("ъ", "j");
            words.Add("ы", "i");
            words.Add("ь", "j");
            words.Add("э", "e");
            words.Add("ю", "yu");
            words.Add("я", "ya");
            words.Add("А", "A");
            words.Add("Б", "B");
            words.Add("В", "V");
            words.Add("Г", "G");
            words.Add("Д", "D");
            words.Add("Е", "E");
            words.Add("Ё", "Yo");
            words.Add("Ж", "Zh");
            words.Add("З", "Z");
            words.Add("И", "I");
            words.Add("Й", "J");
            words.Add("К", "K");
            words.Add("Л", "L");
            words.Add("М", "M");
            words.Add("Н", "N");
            words.Add("О", "O");
            words.Add("П", "P");
            words.Add("Р", "R");
            words.Add("С", "S");
            words.Add("Т", "T");
            words.Add("У", "U");
            words.Add("Ф", "F");
            words.Add("Х", "H");
            words.Add("Ц", "C");
            words.Add("Ч", "Ch");
            words.Add("Ш", "Sh");
            words.Add("Щ", "Sch");
            words.Add("Ъ", "J");
            words.Add("Ы", "I");
            words.Add("Ь", "J");
            words.Add("Э", "E");
            words.Add("Ю", "Yu");
            words.Add("Я", "Ya");

        }

        public static string ToEn(string RU)
        {
            if(string.IsNullOrEmpty(RU)) return RU;
            string EN = RU;
            foreach (KeyValuePair<string, string> pair in words)
                EN = EN.Replace(pair.Key, pair.Value);
            return EN;
        }
    }

    public class Translate
    {
        public static string ToEn(string RU)
        {
            byte[] data = null;
            try
            {
                string url = string.Format("https://translate.googleapis.com/translate_a/single?client=gtx&sl={0}&tl={1}&dt=t&q={2}", "ru", "en", HttpUtility.UrlEncode(RU));
                string outputFile = Path.GetTempFileName();
                using (WebClient wc = new WebClient())
                {
                    wc.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.36");
                    data = wc.DownloadData(url);
                };
            }
            catch { };
            if (data == null) return RU;

            string resp = Encoding.UTF8.GetString(data);
            Regex rx = new Regex("\"[^\"]+\",", RegexOptions.None);
            MatchCollection mc = rx.Matches(resp);
            List<string> subs = new List<string>();
            foreach (Match mx in mc)
                subs.Add(mx.Value.Trim(new char[] { '"', ',', ' ' }));
            string res = "";
            for (int i = 1; i < subs.Count; i++)
                if (RU.Contains(subs[i]))
                    res += subs[i - 1] + " ";
            if (string.IsNullOrEmpty(res)) return RU;
            return res.Trim();
        }
    }
}
