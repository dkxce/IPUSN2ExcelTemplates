//
// Запуск DLL плагина
//

using System;
using System.Drawing;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using PluginsMain;
using System.Xml;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Security;

namespace ExcelTemplatesLib
{
    public class ExcelTemplatesLib : IPluginInterface
    {
        public const string PluginType = "export_documents";
        public const string PluginName = "Вывод документов на основе расширенных шаблонов Excel";
        public const string SupportedDocuments = "счет,акт";
        public const string PluginShortName = "Вывод Excel Templates";

        public string AppDir = null;
        public string CurrDir = XMLSaved<int>.CurrentDirectory();

        public List<string> SupportedDocumentsAsList { get { return new List<string>(SupportedDocuments.Split(',')); }}

        #region From IPluginInterface

        public string GetPluginType() { return PluginType; }
        public string GetPluginName() { return PluginName; }
        public string GetSupportedDocuments() { return SupportedDocuments; }
        public string GetPluginShortName() { return PluginShortName; }        

        public void SetApplicationDirectory(string path)
        {
            if (string.IsNullOrEmpty(path)) return;
            AppDir = path;
        }

        public void SetPluginDirectory(string path)
        {
            if (string.IsNullOrEmpty(path)) return;
            CurrDir = path;
        }

        public void OpenConfig()
        {
            // MessageBox.Show("В этой версии плагина настройки не предусмотрены", PluginName, MessageBoxButtons.OK, MessageBoxIcon.Information);

            CFGForm cfgf = new CFGForm();
            cfgf.Show();
            while (cfgf.IsAlive)
            {
                Application.DoEvents();
                System.Threading.Thread.Sleep(250);
            };
            cfgf.Dispose();
        }

        public void Run(string filePath, string args)
        {
            RunInternal(filePath, args);
        }

        #endregion From IPluginInterface

        private void RunInternal(string filePath, string args)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("Не указан файл для импорта!", PluginName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            };

            if (!File.Exists(filePath))
            {
                MessageBox.Show($"Файл {Path.GetFileName(filePath)} не найден!", PluginName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            };

            string docText = null;
            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    byte[] buffer = new byte[12];
                    fs.Read(buffer, 0, buffer.Length);
                    docText = Encoding.UTF8.GetString(buffer);
                };
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", PluginName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            };

            if (string.IsNullOrEmpty(docText) || (!docText.StartsWith("<?xml")))
            {
                // MessageBox.Show($"Неподдерживаемый формат документа\r\nПоддерживаемый формат: xml", PluginName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProxyQuery(filePath);
                return;
            };

            ExportedDocument ed = null;
            try
            {
                ed = ExportedDocument.Load(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", PluginName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            };

            LoadDocument(ed, filePath);
        }

        private void ProxyQuery(string filePath)
        {
            const string toRun = "excel";
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo(toRun, $"\"{filePath}\"");
                psi.UseShellExecute = true;
                Process.Start(psi);
            }
            catch { };
        }

        private void LoadDocument(ExportedDocument doc, string sourcePath = null)
        {
            if (!SupportedDocumentsAsList.Contains(doc.DocType))
            {
                // MessageBox.Show($"Неподдерживаемый тип документа {doc.DocType}\r\nПоддерживаемый формат: {SupportedDocuments}", PluginName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProxyQuery(sourcePath);
                return;
            };
            
            if (!Initialize())
            {
                MessageBox.Show($"Не удается инициализировать требуемые библиотеки", PluginName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            };

            ProcessDocument(doc, sourcePath);
        }

        private bool Initialize()
        {            
            return InitLibraries() & DeleteOldFiles() & true;
        }

        private bool InitLibraries()
        {
            const string fName = "libpreloader.exe";
            string fPath = Path.Combine(CurrDir, fName);
            string[] files = File.Exists(fPath) ? new string[] { fPath } : Directory.GetFiles(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "libpreloader.exe", SearchOption.AllDirectories);
            foreach (string f in files)
            {
                try
                {
                    ProcessStartInfo psi = new ProcessStartInfo(f);
                    psi.WorkingDirectory = Path.GetDirectoryName(f);
                    psi.CreateNoWindow = true;
                    psi.WindowStyle = ProcessWindowStyle.Hidden;
                    Process proc = Process.Start(psi);
                    proc.WaitForExit();
                    return true;
                }
                catch { };
            };
            return false;
        }

        private void AddFileToOld(string fName)
        {
            string lfName = Path.Combine(CurrDir, "_created_files_list.txt");
            try 
            { 
                using (FileStream fs = new FileStream(lfName, FileMode.Append, FileAccess.Write))
                {
                    byte[] buffer = Encoding.UTF8.GetBytes($"{fName}\r\n");
                    fs.Write(buffer, 0, buffer.Length);
                };
            }
            catch { };
        }

        private bool DeleteOldFiles()
        {
            string lfName = Path.Combine(CurrDir, "_created_files_list.txt");
            try
            {
                List<string> files = new List<string>();
                using (FileStream fs = new FileStream(lfName, FileMode.Open, FileAccess.Read))
                {
                    byte[] buffer = new byte[fs.Length];
                    fs.Read(buffer, 0, buffer.Length);
                    files.AddRange(Encoding.UTF8.GetString(buffer).Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries));
                };
                for(int i=files.Count-1;i>=0;i--)
                {
                    try { File.Delete(files[i]); } catch { };
                    if (!File.Exists(files[i])) files.RemoveAt(i);
                };
                File.Delete(lfName);
                using (FileStream fs = new FileStream(lfName, FileMode.Append, FileAccess.Write))
                {
                    foreach (string f in files)
                    {
                        byte[] buffer = Encoding.UTF8.GetBytes($"{f}\r\n");
                        fs.Write(buffer, 0, buffer.Length);
                    };
                };
            }
            catch { };
            return true;
        }
        
        private string SelectTemplate(string docType)
        {
            string[] files = Directory.GetFiles(Path.Combine(CurrDir, "Templates"), $"*{docType}*.xlsx", SearchOption.TopDirectoryOnly);
            if(files != null && files.Length > 0) 
            {
                string[] values = new string[files.Length];
                for(int i = 0; i < files.Length; i++) values[i] = Path.GetFileName(files[i]);
                int sf = 0;
                InputBox.defWidth = 400;
                InputBox.stayInTop = true;
                DialogResult dr = InputBox.QueryListBox("Формирование документа", $"Выберите шаблон {docType}а из списка:", values, ref sf);
                if (dr != DialogResult.OK) return null;
                return files[sf];
            };
            return null;
        }

        private void ProcessDocument(ExportedDocument doc, string sourcePath = null)
        {
            bool add = string.IsNullOrEmpty(doc.GetDocField("PARTNER_INN")) && string.IsNullOrEmpty(doc.GetDocField("PARTNER_KPP"));
            string tmpName = null;
            string tmpPath = null;

            PluginConfig pc = new PluginConfig();
            if (File.Exists(PluginConfig.FileName)) pc = PluginConfig.Load(PluginConfig.FileName);

            if (pc.StartMode == 3) // ИНН
            {
                string suffix = doc.DocType == "счет" && add ? "_QR" : "";
                foreach (string prefix in new string[] { doc.GetDocField("MYINN") /* Individual Design */ })
                {
                    tmpName = $"_{prefix}_{doc.DocType}{suffix}.xlsx";
                    tmpPath = Path.Combine(Path.Combine(CurrDir, "Templates"), tmpName);
                    if (File.Exists(tmpPath)) break;
                };
            }
            else if (pc.StartMode == 2) // universal
            {
                string suffix = doc.DocType == "счет" && add ? "_QR" : "";
                foreach (string prefix in new string[] { "template" /* Universal Design */ })
                {
                    tmpName = $"_{prefix}_{doc.DocType}{suffix}.xlsx";
                    tmpPath = Path.Combine(Path.Combine(CurrDir, "Templates"), tmpName);
                    if (File.Exists(tmpPath)) break;
                };
            }
            else if (pc.StartMode == 1) // selectable
            {
                tmpPath = SelectTemplate(doc.DocType);
                if (!string.IsNullOrEmpty(tmpPath)) tmpName = Path.GetFileName(tmpPath);
            }
            else if (pc.StartMode == 0) // default
            {
                string suffix = doc.DocType == "счет" && add ? "_QR" : "";
                foreach (string prefix in new string[] { doc.GetDocField("MYINN") /* Individual Design */, "template" /* Universal Design */ })
                {
                    tmpName = $"_{prefix}_{doc.DocType}{suffix}.xlsx";
                    tmpPath = Path.Combine(Path.Combine(CurrDir, "Templates"), tmpName);
                    if (File.Exists(tmpPath)) break;
                };
            };

            if (!File.Exists(tmpPath))
            {
                MessageBox.Show($"Файл шаблона {tmpName} не найден", PluginName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            };

            SmartXLS.WorkBook wb = null;
            try
            {
                wb = new SmartXLS.WorkBook();
                wb.readXLSX(tmpPath);
                wb.Sheet = 0;
                wb.PrintHeader = "";
                wb.PrintFooter = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка открытия шаблона: {ex.Message}", PluginName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            };

            // FILL DOC
            for(int i =0;i<doc.DocFields.Count;i++)
            {
                string id = doc.DocFields[i].id.Trim('%', '$', ' ');
                SetVar(wb, $"%{id}%", doc.DocFields[i].value?.Trim(), false);
            };

            // ENLARGE ITEMS
            if (doc.DocItems.Count > 0)
            {
                EnlargeRows(wb, "%GROUP1%", "%END_GROUP1%", doc.DocItems.Count);
                for (int x = 0; x < doc.DocItems.Count; x++)
                    for(int i = 0; i < doc.DocItems[x].DocFields.Count;i++)
                    {
                        string id = doc.DocItems[x].DocFields[i].id.Trim('%', '$', ' ');
                        SetVar(wb, $"%{id}%", doc.DocItems[x].DocFields[i].value?.Trim(), true);
                    };
            };

            // Add QrCode
            if(doc.DocType == "счет") AddQRCode(doc, wb);
            
            SaveResult(wb, sourcePath);            
        }

        // дублирование строк
        private static void EnlargeRows(SmartXLS.WorkBook wb, string txtFrom, string txtTo, int count)
        {
            int cFrom = FindRow(wb, txtFrom);
            int cTo = FindRow(wb, txtTo);
            if (cFrom > 0 && cTo > 0)
            {
                for (int i = 0; i < count; i++)
                    CopyRowsNext(wb, cFrom, cTo, false);
                DeleteRows(wb, cFrom, cTo, true);
            };
        }

        // установка переменных
        private static void SetVar(SmartXLS.WorkBook wb, string varName, string value, bool onlyFirst)
        {
            for (int c = 0; c <= wb.LastCol; c++)
            {
                for (int r = 0; r <= wb.LastRow; r++)
                {
                    string txt = wb.getText(r, c);
                    if (!txt.Contains("%")) continue;
                    string ntxt = txt.Replace(varName, value);
                    if (ntxt == txt) continue;
                    wb.setText(r, c, ntxt);
                    if (onlyFirst) return;
                };
            };
        }

        // поиск строки с текстом
        private static int FindRow(SmartXLS.WorkBook wb, string text)
        {
            for (int c = 0; c <= wb.LastCol; c++)
            {
                for (int r = 0; r <= wb.LastRow; r++)
                {
                    string txt = wb.getText(r, c);
                    if (txt.Contains(text)) return r;
                };
            };
            return -1;
        }

        // поиск ячейки с текстом
        public static int[] FindRC(SmartXLS.WorkBook wb, string text)
        {
            for (int c = 0; c <= wb.LastCol; c++)
            {
                for (int r = 0; r <= wb.LastRow; r++)
                {
                    string txt = wb.getText(r, c);
                    if (txt.Contains(text)) return new int[] { r, c };
                };
            };
            return null;
        }

        // копируем несколько строк от (rowFrom) до (rowTo) сразу после rowTo
        // withFisrtLast - копировать rowFrom и rowTo, иначе только строки между ними
        public static void CopyRowsNext(SmartXLS.WorkBook wb, int rowFrom, int rowTo, bool withFisrtLast)
        {
            int cols = rowTo - rowFrom - 1 + (withFisrtLast ? 2 : 0);
            wb.insertRange(rowTo + 1, 0, rowTo + cols, wb.LastCol, SmartXLS.WorkBook.ShiftRows);
            wb.copyRange(rowTo + 1, 0, rowTo + cols, wb.LastCol, rowFrom + (withFisrtLast ? 0 : 1), 0, rowTo - (withFisrtLast ? 0 : 1), wb.LastCol);
        }

        // удаляем строки от (rowFrom) до (rowTo)
        // withFisrtLast - удаляем rowFrom и rowTo, иначе только строки между ними
        public static void DeleteRows(SmartXLS.WorkBook wb, int rowFrom, int rowTo, bool withFisrtLast)
        {
            if (withFisrtLast)
                wb.deleteRange(rowFrom, 0, rowTo, wb.LastCol, SmartXLS.WorkBook.ShiftRows);
            else
                wb.deleteRange(rowFrom + 1, 0, rowTo - 1, wb.LastCol, SmartXLS.WorkBook.ShiftRows);
        }

        // добавляем QRCode
        private static void AddQRCode(ExportedDocument doc, SmartXLS.WorkBook wb)
        {
            string contract = doc.GetDocField("BYCONTRACT");
            if (!string.IsNullOrEmpty(contract)) contract = "по договору " + contract.ToLower().Replace("по договору", "").Trim();

            string sum_str = (new Regex(@"[\d\s,.]+", RegexOptions.IgnoreCase)).Match(doc.GetDocField("RESULT")).Groups[0].Value;
            sum_str = sum_str.Replace(" ", "").Replace(",", ".");
            double.TryParse(sum_str, System.Globalization.NumberStyles.AllowDecimalPoint, System.Globalization.CultureInfo.InvariantCulture, out double sum_dbl);

            // ST0001
            string QR_TEXT = String.Format(
                    "ST00011|Name={0}|PersonalAcc={1}|BankName={2}|BIC={3}|CorrespAcc={4}|PayeeINN={5}|Sum={6}|Purpose={7}",
                    /* 0 */ doc.GetDocField("MYCOMPANY") /* Получатель */,
                    /* 1 */ doc.GetDocField("MYACCOUNT") /* расч счет  */,
                    /* 2 */ doc.GetDocField("MYBANK")    /* банк       */,
                    /* 3 */ doc.GetDocField("MYBIK")     /* БИК        */,
                    /* 4 */ doc.GetDocField("MYKS")      /* корр счета */,
                    /* 5 */ doc.GetDocField("MYINN")     /* ИНН        */,
                    /* 6 */ (int)(sum_dbl*100)           /* в копейках */,
                    /* 7 */ String.Format("Оплата счета #{0} от {1} {2}", doc.GetDocField("NUMBER"), doc.GetDocField("DATE"), contract)
                    );

            bool add = string.IsNullOrEmpty(doc.GetDocField("PARTNER_INN")) && string.IsNullOrEmpty(doc.GetDocField("PARTNER_KPP"));

            int[] rcf = FindRC(wb, "%IMAGEFROM%");
            int[] rct = FindRC(wb, "%IMAGETO%");
            if (rcf != null) wb.setText(rcf[0], rcf[1], "");
            if (rct != null) wb.setText(rct[0], rct[1], "");
            if (add)
            {
                if ((rcf != null) && (rct != null))
                    AddQrCode(wb, rcf[0], rcf[1], rct[0] + 1, rct[1] + 1, QR_TEXT);
                else if (rcf != null)
                    AddQrCode(wb, rcf[0], rcf[1], -1, -1, QR_TEXT);
            };
        }

        // добавляем QRCode
        private static void AddQrCode(SmartXLS.WorkBook wb, int frow, int fcol, int trow, int tcol, string text)
        {
            if (frow < 0) return;
            if (fcol < 0) return;
            
            ThoughtWorks.QRCode.Codec.QRCodeEncoder qr = new ThoughtWorks.QRCode.Codec.QRCodeEncoder();
            qr.QRCodeEncodeMode = ThoughtWorks.QRCode.Codec.QRCodeEncoder.ENCODE_MODE.BYTE;
            qr.QRCodeVersion = 0;
            qr.QRCodeScale = 3;

            Bitmap bmpI = qr.Encode(text, Encoding.UTF8);
            Bitmap bmpO = new System.Drawing.Bitmap(bmpI.Width + 8, bmpI.Width + 8);
            Graphics g = Graphics.FromImage(bmpO);
            g.Clear(System.Drawing.Color.White);
            g.DrawImage(bmpI, new System.Drawing.Point(4, 4));
            g.Dispose();

            string tmpF = Path.GetTempFileName();
            bmpO.Save(tmpF);
            wb.addPicture(fcol, frow, tcol, trow, tmpF);
            File.Delete(tmpF);
        }

        // Сохраняем в файл
        private void SaveResult(SmartXLS.WorkBook wb, string fName = null) 
        {
            bool add2old = true;
            if (string.IsNullOrEmpty(fName))
            {
                fName = Path.GetTempFileName();
                if (File.Exists(fName)) File.Delete(fName);
                fName += ".xlsx";                
            }
            else
            {
                string fExt = Path.GetExtension(fName).ToLower();
                if(fExt != ".xlsx") fName += ".xlsx";
                add2old = false;
            };
            wb.writeXLSX(fName);            
            ProxyQuery(fName);
            if(add2old) AddFileToOld(fName);
        }
    }
}
