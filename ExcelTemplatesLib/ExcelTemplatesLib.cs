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
using System.Drawing.Imaging;
using System.Security.Cryptography;

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

        private string SelectTemplateLast(string docType, PluginConfig pc)
        {
            for (int i = pc.LastTemplates.Count - 1; i >= 0; i--)
                if (pc.LastTemplates[i].d == docType)
                    return Path.Combine(Path.Combine(CurrDir, "Templates"), pc.LastTemplates[i].f);
            return null;
        }

        private string SelectTemplate(string docType, PluginConfig pc)
        {
            List<string> fls = new List<string>();
            fls.AddRange(Directory.GetFiles(Path.Combine(CurrDir, "Templates"), $"*{docType}*.xlsx", SearchOption.TopDirectoryOnly));
            fls.AddRange(Directory.GetFiles(Path.Combine(CurrDir, "Templates"), $"*{docType}*.xlsm", SearchOption.TopDirectoryOnly));
            string[] files = fls.ToArray();
            if (files != null && files.Length > 0) 
            {
                List<string> values = new List<string>();
                for(int i = 0; i < files.Length; i++) values.Add(Path.GetFileName(files[i]));
                int sf = 0;
                for(int i= pc.LastTemplates.Count-1; i>=0;i--)
                    if (pc.LastTemplates[i].d == docType)
                    {
                        sf = values.IndexOf(pc.LastTemplates[i].f);
                        if (sf < 0) sf = 0;
                        pc.LastTemplates.RemoveAt(i);
                        break;
                    };
                InputBox.defWidth = 400;
                InputBox.stayInTop = true;
                DialogResult dr = InputBox.QueryListBox("Формирование документа", $"Выберите шаблон {docType}а из списка:", values.ToArray(), ref sf);
                if (dr != DialogResult.OK) return null;
                pc.LastTemplates.Add(new PluginConfig.DocumentFile(docType, values[sf]));
                PluginConfig.SaveHere("config.xml", pc);
                return files[sf];
            };
            return null;
        }

        private void ProcessDocument(ExportedDocument doc, string sourcePath = null)
        {
            PluginConfig pc = new PluginConfig();
            if (File.Exists(PluginConfig.FileName)) pc = PluginConfig.Load(PluginConfig.FileName);

            bool add = string.IsNullOrEmpty(doc.GetDocField("PARTNER_INN")) && string.IsNullOrEmpty(doc.GetDocField("PARTNER_KPP"));
            if ((!add) && pc.QRIP && doc.GetDocField("PARTNER").ToLower().Contains("индивидуальный предприниматель")) add = true;

            string tmpName = null;
            string tmpPath = null;            

            if(pc.StartMode == 4) // Last Selected
            {
                tmpPath = SelectTemplateLast(doc.DocType, pc);
                if (!string.IsNullOrEmpty(tmpPath)) tmpName = Path.GetFileName(tmpPath);
            }
            else if (pc.StartMode == 3) // ИНН
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
                tmpPath = SelectTemplate(doc.DocType, pc);
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
            bool isXLSM = false;
            try
            {
                wb = new SmartXLS.WorkBook();
                wb.readXLSX(tmpPath);
                wb.Sheet = 0;
                wb.PrintHeader = "";
                wb.PrintFooter = "";
                isXLSM = Path.GetExtension(tmpPath).ToLower().EndsWith("m");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка открытия шаблона: {ex.Message}", PluginName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            };

            // Fill Defaults
            SetVar(wb, "%TODAY%", DateTime.Now.ToString("dd.MM.yyyy"), false);
            SetVar(wb, "%NOW%", DateTime.Now.ToString("HH:mm:ss"), false);
            SetVar(wb, "%SOFTWARE%", "Excel Template плагин для ИП УСН 2", false);

            // FILL DOC            
            string dInfo = $"TYPE={doc.DocType}";
            string dNum = "";
            string dDat = "";
            for (int i =0;i<doc.DocFields.Count;i++)
            {
                string id = doc.DocFields[i].id.Trim('%', '$', ' ');
                SetVar(wb, $"%{id}%", doc.DocFields[i].value?.Trim(), false);
                SetVarEn(wb, $"%{id}_EN%", doc.DocFields[i].value?.Trim(), false);
                if (id == "NUMBER") dInfo += $"|{id}=" + doc.DocFields[i].value?.Trim();
                if (id == "SHORTDATE") dInfo += $"|{id}=" + doc.DocFields[i].value?.Trim();
                if (id == "NUMBER") dNum = doc.DocFields[i].value?.Trim().Replace("$","*");
                if (id == "SHORTDATE") dDat = doc.DocFields[i].value?.Trim().Replace("$", "*");
            };
            string matrix = System.Convert.ToBase64String(System.Text.Encoding.GetEncoding(1251).GetBytes(dInfo));
            string barcod = $"${doc.IntType}${System.Web.HttpUtility.UrlEncode(dNum)}${System.Web.HttpUtility.UrlEncode(dDat)}";
            barcod = Regex.Replace(barcod.ToUpper(), "[^0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ\\-\\.\\s\\$\\/\\+\\%\\*]", " ");

            // ENLARGE ITEMS
            if (doc.DocItems.Count > 0)
            {
                EnlargeRows(wb, "%GROUP1%", "%END_GROUP1%", doc.DocItems.Count);
                for (int x = 0; x < doc.DocItems.Count; x++)
                    for(int i = 0; i < doc.DocItems[x].DocFields.Count;i++)
                    {
                        string id = doc.DocItems[x].DocFields[i].id.Trim('%', '$', ' ');
                        SetVar(wb, $"%{id}%", doc.DocItems[x].DocFields[i].value?.Trim(), true);
                        SetVarEn(wb, $"%{id}_EN%", doc.DocItems[x].DocFields[i].value?.Trim(), true);
                    };
            };

            // Codes
            {
                int _r = -1; int _c = -1; bool _repl = true;

                // Add QrCode
                if (doc.DocType == "счет") AddQRCode(doc, wb, pc);

                // Add Code39
                if (_repl) { _r = -1; _c = -1; };
                AddBarCode(wb, barcod, pc.Code39Bar, pc.SingleCode, ref _r, ref _c, out _repl);

                // Add Matrix Code
                if (_repl) { _r = -1; _c = -1; };
                AddMatrixCode(wb, matrix, pc.MatrixBar, pc.MatrixCode, ref _r, ref _c, out _repl);                
            };

            SaveResult(wb, sourcePath, isXLSM);            
        }

        // дублирование строк
        private static void EnlargeRows(SmartXLS.WorkBook wb, string txtFrom, string txtTo, int count)
        {
            List<SmartXLS.ShapePos> poses = new List<SmartXLS.ShapePos>();
            for (int i = 0; i < wb.PictureCount; i++)
                poses.Add(wb.getPictureShape(i).ShapePos);

            int addedRows = wb.LastRow;

            int cFrom = FindRow(wb, txtFrom);
            int cTo = FindRow(wb, txtTo);
            if (cFrom > 0 && cTo > 0)
            {
                for (int i = 0; i < count; i++)
                    CopyRowsNext(wb, cFrom, cTo, false);
                DeleteRows(wb, cFrom, cTo, true);
            };

            addedRows = wb.LastRow - addedRows;
            for (int i = 0; i < poses.Count; i++)
                if((poses[i].Y1 > cTo) && (poses[i].Y1 + addedRows > 0) && (poses[i].Y2 + addedRows > 0))
                    wb.getPictureShape(i).setPosition(poses[i].X1, poses[i].Y1 + addedRows, poses[i].X2, poses[i].Y2 + addedRows);
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
                    try { wb.setText(r, c, ntxt); } catch (Exception ex) { wb.setText(r, c, txt); };
                    if (onlyFirst) return;
                };
            };
        }

        private static void SetVarEn(SmartXLS.WorkBook wb, string varName, string value, bool onlyFirst)
        {
            for (int c = 0; c <= wb.LastCol; c++)
            {
                for (int r = 0; r <= wb.LastRow; r++)
                {
                    string txt = wb.getText(r, c);
                    if (!txt.Contains(varName)) continue;
                    string ntxt = txt.Replace(varName, Translit(value));
                    if (ntxt == txt) continue;
                    try { wb.setText(r, c, ntxt); } catch (Exception ex) { wb.setText(r, c, txt); };
                    if (onlyFirst) return;
                };
            };
        }

        private static bool FindVar(SmartXLS.WorkBook wb, string varName)
        {
            for (int c = 0; c <= wb.LastCol; c++)
                for (int r = 0; r <= wb.LastRow; r++)
                    if (wb.getText(r, c).Contains(varName)) return true;
            return false;
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
        private static void AddQRCode(ExportedDocument doc, SmartXLS.WorkBook wb, PluginConfig pc)
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
            if ((!add) && pc.QRIP && doc.GetDocField("PARTNER").ToLower().Contains("индивидуальный предприниматель")) add = true;

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

        private void AddMatrixCode(SmartXLS.WorkBook wb, string data, bool paste, string codebar, ref int _r, ref int _c, out bool replaced)
        {
            replaced = false;
            int ir = -1;
            int ic = -1;
            for (int c = 0; c <= wb.LastCol; c++)
                for (int r = 0; r <= wb.LastRow; r++)
                    if(wb.getText(r,c).Contains("%MATRIX%"))
                    {
                        wb.setText(r, c, wb.getText(r, c).Replace("%MATRIX%", ""));
                        ir = r;
                        ic = c;
                        replaced = true;
                        _r = -1;
                        _c = -1;
                    };

            if (!paste) return;
                        
            DataMatrix.net.DmtxImageEncoder ie = new DataMatrix.net.DmtxImageEncoder();
            DataMatrix.net.DmtxImageEncoderOptions ops = new DataMatrix.net.DmtxImageEncoderOptions();
            ops.ModuleSize = 3;
            Bitmap res = ie.EncodeImage(data, ops);            
            
            string zFile = Path.Combine(CurrDir, "zint.exe");
            if (codebar != "Data Matrix" && File.Exists(zFile))
            {
                codebar = codebar.Replace("-", "").Replace(" ", "").Trim().ToUpper();
                string outFile = Path.Combine(CurrDir, "out2.png");
                string cmdLine = $"-b {codebar} -o \"{outFile}\" -d \"{data}\"";
                
                if (File.Exists(outFile)) try { File.Delete(outFile); } catch { };
                ProcessStartInfo psi = new ProcessStartInfo(zFile, cmdLine);
                psi.WindowStyle = ProcessWindowStyle.Hidden;
                psi.CreateNoWindow = true;
                try
                {
                    Process proc = Process.Start(psi);
                    proc.WaitForExit(5000);
                    if (File.Exists(outFile))
                    {
                        res = (Bitmap)Bitmap.FromFile(outFile);
                        res = ResizeImage(res, 116, 116);
                    };
                }
                catch { };
            };

            int reswi = res.Width;
            byte[] bytes = GetImageAsBytes(res);

            if (_r >= 0) { ir = _r; replaced = false; };
            if (_c >= 0) { ic = _c; replaced = false; };
            if (ir < 0 || ic < 0)
            {
                ir = 0;
                ic = 0;
                int mrh = 0;
                int pph = 16838 /* A4 */; // wb.PrintPaperHeight;
                while (mrh < pph) mrh += wb.getRowHeight(ir++);
            };            

            wb.addPicture(_c = ic, _r = ir, ic, ir, bytes);
        }

        private void AddBarCode(SmartXLS.WorkBook wb, string data, bool paste, string codebar, ref int _r, ref int _c, out bool replaced)
        {
            replaced = false;
            int ir = -1;
            int ic = -1;
            for (int c = 0; c <= wb.LastCol; c++)
                for (int r = 0; r <= wb.LastRow; r++)
                    if (wb.getText(r, c).Contains("%BARCODE%"))
                    {
                        wb.setText(r, c, wb.getText(r, c).Replace("%BARCODE%", ""));
                        ir = r;
                        ic = c;
                        replaced = true;
                        _r = -1;
                        _c = -1;
                    };

            if (!paste) return;

            if (_r >= 0) { ir = _r; replaced = false; };
            if (_c >= 0) { ic = _c; replaced = false; };
            if (ir < 0 || ic < 0)
            {
                ir = 0;
                ic = 0;
                int mrh = 0;
                int pph = 16838 /* A4 */; // wb.PrintPaperHeight;
                while (mrh < pph) mrh += wb.getRowHeight(ir++);
            };

            DSBarCode.BarCodeCtrl bc = new DSBarCode.BarCodeCtrl();
            bc.ShowFooter = false;
            bc.ShowHeader = false;
            bc.BarCodeHeight = 40;
            bc.Width = 940;
            bc.Height = 60;
            bc.BarCode = data;

            Bitmap res = null;
            string zFile = Path.Combine(CurrDir, "zint.exe");
            if ((codebar != "Code 39" || (!bc.IsValid)) && File.Exists(zFile))
            {
                codebar = codebar.Replace("-", "").Replace(" ", "").Trim().ToUpper();
                string outFile = Path.Combine(CurrDir, "out1.png");
                string cmdLine = $"-b {codebar} -o \"{outFile}\" -d \"{data}\"";

                if (File.Exists(outFile)) try { File.Delete(outFile); } catch { };
                ProcessStartInfo psi = new ProcessStartInfo(zFile, cmdLine);
                psi.WindowStyle = ProcessWindowStyle.Hidden;
                psi.CreateNoWindow = true;
                try
                {
                    Process proc = Process.Start(psi);
                    proc.WaitForExit(5000);
                    if (File.Exists(outFile))
                    {
                        res = (Bitmap)Bitmap.FromFile(outFile);
                        res = CropImage(res, 960, 60);
                    };
                }
                catch { };
            };
            if (res != null)
            {
                string tmpF = Path.GetTempFileName();
                res.Save(tmpF);
                res.Dispose();
                wb.addPicture(_c = ic, _r = ir, ic, ir, tmpF);
                File.Delete(tmpF);
            }
            else if (bc.IsValid)
            {
                string tmpF = Path.GetTempFileName();
                bc.SaveImage(tmpF);
                wb.addPicture(_c = ic, _r = ir, ic, ir, tmpF);
                File.Delete(tmpF);                
            };
        }

        public static Bitmap ResizeImage(Image image, int width, int height, bool keepAspectRation = true)
        {
            int sourceWidth = image.Width;
            int sourceHeight = image.Height;
            int sourceX = 0;
            int sourceY = 0;
            int destX = 0;
            int destY = 0;

            int destWidth = width;
            int destHeight = height;

            if (keepAspectRation)
            {
                float nPercent = 0;
                float nPercentW = ((float)width / (float)sourceWidth);
                float nPercentH = ((float)height / (float)sourceHeight);

                if (nPercentH < nPercentW)
                {
                    nPercent = nPercentH;
                    destX = System.Convert.ToInt16((width - (sourceWidth * nPercent)) / 2);
                }
                else
                {
                    nPercent = nPercentW;
                    destY = System.Convert.ToInt16((height - (sourceHeight * nPercent)) / 2);
                };

                destWidth = (int)(sourceWidth * nPercent);
                destHeight = (int)(sourceHeight * nPercent);
            };

            Bitmap bmPhoto = new Bitmap(width, height);//, PixelFormat.Format24bppRgb);
            //bmPhoto.SetResolution(image.HorizontalResolution, image.VerticalResolution);
            Graphics g = Graphics.FromImage(bmPhoto);
            //g.Clear(Color.White);
            //g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.DrawImage(image, new Rectangle(destX, destY, destWidth, destHeight), new Rectangle(sourceX, sourceY, sourceWidth, sourceHeight), GraphicsUnit.Pixel);
            g.Dispose();
            return bmPhoto;
        }

        public static Bitmap CropImage(Image image, int width, int height)
        {            
            Bitmap bmPhoto = new Bitmap(width, height);
            Graphics g = Graphics.FromImage(bmPhoto);
            g.DrawImage(image, width / 2 - image.Width / 2, height / 2 - image.Height / 2);
            g.Dispose();
            return bmPhoto;
        }

        private byte[] GetImageAsBytes(Bitmap bmp)
        {
            MemoryStream ms = new MemoryStream();
            bmp.Save(ms, ImageFormat.Png);
            byte[] bmpBytes = ms.GetBuffer();
            bmp.Dispose();
            ms.Close();
            return bmpBytes;
        }

        // Сохраняем в файл
        private void SaveResult(SmartXLS.WorkBook wb, string fName = null, bool isXLSM = false) 
        {
            bool add2old = true;
            if (string.IsNullOrEmpty(fName))
            {
                fName = Path.GetTempFileName();
                if (File.Exists(fName)) File.Delete(fName);
                fName += isXLSM ? ".xlsm" : ".xlsx";
            }
            else
            {
                string fExt = Path.GetExtension(fName).ToLower();
                if (isXLSM)
                {
                    if (fExt != ".xlsm") fName += ".xlsm";
                }
                else
                {
                    if (fExt != ".xlsx") fName += ".xlsx";
                };
                add2old = false;
            };
            wb.writeXLSX(fName);            
            ProxyQuery(fName);
            if(add2old) AddFileToOld(fName);
        }

        private static string Translit(string text, bool translate = true)
        {
            string res = TRANSLIT.Translit.ToEn(text);
            if ((!translate) || (text == res)) return res;
            return TRANSLIT.Translate.ToEn(text);
        }
    }
}
