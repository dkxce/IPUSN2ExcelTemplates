using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace System
{
    /// <summary>
    ///     Windows Console Class
    /// </summary>
    public static class WinConsoleApplication
    {
        #region WinAPI

        private const UInt32 GENERIC_WRITE = 0x40000000;
        private const UInt32 GENERIC_READ = 0x80000000;
        private const UInt32 FILE_SHARE_READ = 0x00000001;
        private const UInt32 FILE_SHARE_WRITE = 0x00000002;
        private const UInt32 OPEN_EXISTING = 0x00000003;
        private const UInt32 FILE_ATTRIBUTE_NORMAL = 0x80;
        private const UInt32 ERROR_ACCESS_DENIED = 5;
        private const UInt32 ATTACH_PARRENT = 0xFFFFFFFF;

        [DllImport("kernel32.dll", EntryPoint = "AllocConsole", SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        private static extern int AllocConsole();

        [DllImport("kernel32.dll", EntryPoint = "AttachConsole", SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        private static extern UInt32 AttachConsole(UInt32 dwProcessId);

        [DllImport("kernel32.dll", EntryPoint = "CreateFileW", SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        private static extern IntPtr CreateFileW(string lpFileName, UInt32 dwDesiredAccess, UInt32 dwShareMode, IntPtr lpSecurityAttributes, UInt32 dwCreationDisposition, UInt32 dwFlagsAndAttributes, IntPtr hTemplateFile);

        [DllImport("kernel32.dll", EntryPoint = "FreeConsole", SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        private static extern bool FreeConsole();

        #endregion

        /// <summary>
        ///     Initialization Status (0 - not init, 1 - simple, 2 - attached)
        /// </summary>
        private static byte InitializationStatus = 0;

        /// <summary>
        ///     Default Console CodePage
        /// </summary>
        public static Encoding ConsoleEncoding = Encoding.GetEncoding(866);

        #region Public Init Methods

        /// <summary>
        ///     Simple Console Init
        /// </summary>
        public static void Initialize()
        {
            if (InitializationStatus != 0) return;
            InitializationStatus = 1; // simple
            AllocConsole();
            try
            {
                Console.OutputEncoding = ConsoleEncoding;
                Console.InputEncoding = ConsoleEncoding;
            }
            catch { };
        }

        /// <summary>
        ///     Initialize Standard Console (with new window if need)
        /// </summary>
        /// <param name="createNewConsole"></param>
        public static void Initialize(bool createNewConsole)
        {
            Initialize(createNewConsole, true, false);
        }

        /// <summary>
        ///     Initialize Standard Console (with new window if need)
        /// </summary>
        /// <param name="createNewConsole"></param>
        /// <param name="allowWrite"></param>
        public static void Initialize(bool createNewConsole, bool redirectStdOut)
        {
            Initialize(createNewConsole, redirectStdOut, false);
        }

        /// <summary>
        ///     Initialize Standard Console (with new window if need)
        /// </summary>
        /// <param name="createNewConsole"></param>
        /// <param name="allowWrite"></param>
        /// <param name="allowRead"></param>
        public static void Initialize(bool createNewConsole, bool redirectStdOut, bool redirectStdIn)
        {
            if (InitializationStatus != 0) return;
            InitializationStatus = 2; // Attached

            bool consoleAttached = true;
            if (createNewConsole || ((AttachConsole(ATTACH_PARRENT) == 0) && (Marshal.GetLastWin32Error() != ERROR_ACCESS_DENIED)))
                consoleAttached = AllocConsole() != 0;

            if (consoleAttached)
            {
                if (redirectStdOut) InitStdOut();
                if (redirectStdIn) InitStdIn();
            };
            try
            {
                Console.OutputEncoding = ConsoleEncoding;
                Console.InputEncoding = ConsoleEncoding;
            }
            catch { };
        }

        /// <summary>
        ///     Deinit Console
        /// </summary>
        public static void DeInitialize()
        {
            if (InitializationStatus == 2)
            {
                Console.Out.Close();
                Console.In.Close();
                FreeConsole();
            };
            InitializationStatus = 0; // Not Init
        }

        #endregion Public Init Methods

        #region Private Init Methods

        /// <summary>
        ///     Init/ReInit StdOut
        /// </summary>
        private static void InitStdOut()
        {
            FileStream fs = CreateFileStream("CONOUT$", GENERIC_WRITE, FILE_SHARE_WRITE, FileAccess.Write);
            if (fs != null)
            {
                StreamWriter writer = new StreamWriter(fs, ConsoleEncoding);
                writer.AutoFlush = true;
                Console.SetOut(writer);
                Console.SetError(writer);
            };
        }

        /// <summary>
        ///     Init/ReInit StdIn
        /// </summary>
        private static void InitStdIn()
        {
            FileStream fs = CreateFileStream("CONIN$", GENERIC_READ, FILE_SHARE_READ, FileAccess.Read);
            if (fs != null)
            {
                Console.SetIn(new StreamReader(fs, ConsoleEncoding));
            };
        }

        private static FileStream CreateFileStream(string name, uint win32DesiredAccess, uint win32ShareMode, FileAccess dotNetFileAccess)
        {
            Microsoft.Win32.SafeHandles.SafeFileHandle file = new Microsoft.Win32.SafeHandles.SafeFileHandle(CreateFileW(name, win32DesiredAccess, win32ShareMode, IntPtr.Zero, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, IntPtr.Zero), true);
            if (!file.IsInvalid)
            {
                FileStream fs = new FileStream(file, dotNetFileAccess);
                return fs;
            };
            return null;
        }

        #endregion Private Init Methods        
    }
}
