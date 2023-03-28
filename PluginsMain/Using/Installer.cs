using System;
using System.Xml;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Security.Policy;
using MSol.PluginsMain;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace PluginsMain
{
    public class Installer
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool TerminateProcess(IntPtr hProcess, uint uExitCode);

        public static void Install()
        {
            string path = @"C:\IPUSN2";
            if(!AskToKill(ref path)) return;
            if (!Directory.Exists(path)) path = GetAppPath();
            if (!Directory.Exists(path))
            {
                InputBox.defWidth = 400;
                if (InputBox.QueryDirectoryBox("Установка плагина для ИП УСН2", "Выберите папку с программой:", ref path) != DialogResult.OK)
                {
                    MessageBox.Show("Установка прервана пользователем", "Установка плагина для ИП УСН2", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                };
            };
            if (!Directory.Exists(path))
            {
                MessageBox.Show("Папка с программой не найдена", "Установка плагина для ИП УСН2", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            };
            path = path.Trim('\\') + @"\";
            Console.WriteLine("Установка...");            
            RewriteIni(Path.Combine(path, "config.ini"));
            RewriteTemplates(path);
            CopyFiles(path);
            Console.WriteLine($"Установка успешно завершена");
            if(MessageBox.Show("Установка успешно завершена\r\n\r\nОткрыть файлы шаблонов для редактирования?", "Установка плагина для ИП УСН2", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                ProxyExcel(
                    Path.Combine(path, @"Plugins\ExcelTemplate\Templates\_template_акт.xlsx"),
                    Path.Combine(path, @"Plugins\ExcelTemplate\Templates\_template_счет.xlsx"),
                    Path.Combine(path, @"Plugins\ExcelTemplate\Templates\_template_счет_QR.xlsx")
                    );
        }

        private static void ProxyExcel(params string[] filePath)
        {
            const string toRun = "excel";
            try
            {
                string args = "";
                foreach (string f in filePath) args += (args.Length > 0 ? " " : "") + $"\"{f}\"";
                ProcessStartInfo psi = new ProcessStartInfo(toRun, args);
                psi.UseShellExecute = true;
                Process.Start(psi);
            }
            catch { };
        }

        private static string GetAppPath()
        {
            List<string> w2f = new List<string>();
            w2f.Add(Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            w2f.Add(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu));
            w2f.Add(@"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\");

            foreach (string w in w2f)
                try
                {
                    foreach (string f in GetFiles(w, "*УСН*.lnk"))
                    {
                        ShellLink sl = ShellLink.LoadFromFile(f);
                        return Path.GetDirectoryName(sl.TargetPath);
                    };
                }
                catch (Exception ex) { };
            return null;
        }

        private static List<string> GetFiles(string path, string pattern)
        {
            List<string> files = new List<string>();
            string[] directories = new string[] { };

            try
            {
                files.AddRange(Directory.GetFiles(path, pattern, SearchOption.TopDirectoryOnly));
                directories = Directory.GetDirectories(path);
            }
            catch (UnauthorizedAccessException) { };

            foreach (var directory in directories)
                try { files.AddRange(GetFiles(directory, pattern)); }
                catch (UnauthorizedAccessException) { };

            return files;
        }

        private static void RewriteIni(string filePath)
        {
            List<string> lines = new List<string>();
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                StreamReader sr = new StreamReader(fs, Encoding.GetEncoding(1251));
                while (!sr.EndOfStream)
                    lines.Add(sr.ReadLine());
            };
            for (int i = 0; i < lines.Count; i++)
                if (lines[i].StartsWith("excel ="))
                {
                    string p = Path.Combine(Path.GetDirectoryName(filePath), @"Plugins\ExcelTemplate\ExcelTemplates.exe");
                    lines[i] = $"excel = {p}";
                    break;
                };
            Console.WriteLine(" Файл конфигурации успешно пропатчен");
        }

        private static void RewriteTemplates(string toPath)
        {
            string p = Path.Combine(Path.GetDirectoryName(toPath), @"Templates\");
            Directory.CreateDirectory(p);
            string subs = Path.Combine(XMLSaved<int>.CurrentDirectory(), @"Templates\ToReplace");
            string[] files = Directory.GetFiles(subs, "*.*", SearchOption.AllDirectories);
            foreach(string f in files)
            {
                string fs = f.Substring(subs.Length).Trim('\\');
                string fd = Path.Combine(p, fs);                
                File.Copy(f, fd, true);
                Console.WriteLine($" Файл шаблона {fs} успешно скопирован");
            };
        }

        private static void CopyFiles(string toPath)
        {
            string p = Path.Combine(Path.GetDirectoryName(toPath), @"Plugins\ExcelTemplate\");
            Directory.CreateDirectory(p);
            string subs = XMLSaved<int>.CurrentDirectory();
            string[] files = Directory.GetFiles(subs, "*.*", SearchOption.AllDirectories);
            foreach(string f in files)
            {
                string fs = f.Substring(subs.Length).Trim('\\');
                string fd = Path.Combine(p, fs);
                Directory.CreateDirectory(Path.GetDirectoryName(fd));
                string sfs = Path.GetFileName(fs);
                if (Path.GetExtension(sfs).ToLower() == ".xlsx" && File.Exists(fd))
                    if (MessageBox.Show($"Переписать существующий шаблон `{sfs}`?", "Установка плагина", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        continue;
                File.Copy(f, fd, true);
                Console.WriteLine($" Файл {fs} успешно скопирован");
            };
        }

        public static bool AskToKill(ref string path)
        {
            while (true)
            {
                Process[] procs = Process.GetProcessesByName("IPUSN2");
                if (procs == null || procs.Length == 0) return true;
                path = Path.GetDirectoryName(procs[0].MainModule.FileName);
                DialogResult dr = MessageBox.Show("Программа ИП УСН2 запущена!\r\nЗакройте программу!", "Установка плагина для ИП УСН2", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);
                if (dr == DialogResult.Abort) return false;
                if (dr == DialogResult.Ignore)
                {
                    foreach (Process proc in procs)
                    {
                        try { TerminateProcess(proc.Handle, 0); } catch { };
                        try { proc.Kill(); } catch { };
                    };
                };
                System.Threading.Thread.Sleep(1000);
            };
        }
    }
}
