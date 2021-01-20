using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using Microsoft.Win32;
using System.Threading;
using System.Reflection;
using System.DirectoryServices;
using System.Collections;
using System.Management;
using System.Net.Mail;
using System.Windows;
using System.Security.Principal;
using System.DirectoryServices.ActiveDirectory;
using System.DirectoryServices.AccountManagement;
using System.Data;
using System.Configuration;
using System.Text.RegularExpressions;

namespace Install_CK11
{
    #region    class OSInfo
    class OSInfo
    {
        public string Caption;
        public string OSArchitecture;
        public string Version;
        public string OSLanguage;
        public OSInfo(string Caption, string OSArchitecture,String Version, string OSLanguage)
        {
            this.Caption = Caption;
            this.OSArchitecture = OSArchitecture;
            this.Version = Version;
            this.OSLanguage = OSLanguage;
        }
            public OSInfo()
        {
            this.Caption = null;
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
                foreach (ManagementObject os in searcher.Get())
                {
                    this.Caption = os["Caption"].ToString();
                    this.OSArchitecture=os["OSArchitecture"].ToString();
                    this.OSLanguage=os["OSLanguage"].ToString();
                    this.Version=os["Version"].ToString();
                    break;
                    //COMPUTER : AK47
                    //CLASS: ROOT\CIMV2: Win32_OperatingSystem
                    //QUERY    : SELECT* FROM Win32_OperatingSystem
                    //Caption: Майкрософт Windows 10 Корпоративная LTSC
                    //CSName: AK47
                    //OSArchitecture: 64-разрядная
                    //OSLanguage: 1049
                    //SystemDirectory: C:\Windows\system32
                    //SystemDrive: C:
                    //TotalSwapSpaceSize:
                    //TotalVirtualMemorySize: 22208168
                    //TotalVisibleMemorySize: 16703144
                    //Version: 10.0.17763
                    //WindowsDirectory: C:\Windows
                    //BuildNumber	17763
                    //Codeset 1251
                    //CountryCode 7
                    //Locale 0419
                }
            }
            catch { }
        }
        public bool Compare(OSInfo OS)
        {
            bool Compare = false;
            if(this.OSLanguage.Equals(OS.OSLanguage))
                if (String.Compare(this.OSArchitecture, OS.OSArchitecture, false) >= 0)
                    if (String.Compare(this.Version,OS.Version,false)>=0)
                        if (this.Version.Contains(OS.Version))
                            Compare = true;
            return Compare;
        }
        public string Info()
        {
            string Info = this.Caption + " " + this.OSArchitecture + " " + this.Version +" "+this.OSLanguage;
            return Info;
        }
    }
    #endregion
    class Program
    {
        [DllImport("Kernel32.dll")]
        private static extern IntPtr GetConsoleWindow();
        [DllImport("User32.dll")]
        private static extern bool ShowWindow(IntPtr hwind, int cmdShow);

        public static string __Error;
        const string __root_CIMv2 = @"root\CIMV2";
        #region Config 
        static string Distrib_Folder = @"\\fs2-oduyu\CK2007\СК-11\";
        static string Distrib_Folder_Runtime = @"Runtimes";
        static string Distrib_Folder_CK11 = @"Дистрибутив клиента  СК-11";
        static string autoinstaller_CK11 = "AutoInstall CK11.exe";
        static string installer_CK11 = "SetupClient.exe";
        static string Service_User = "Svc-ck11cl-oduyu";
        static string Service_domain = "oduyu";
        static string MailServer = "";
        static string emailOikAdmin = "";
        #endregion
        static string Hostname;
        static string LocalAdministratorsGroup = null;
        static ulong minPhysicalMemory = 6144;
        const char hr_char = '─';
        static int hr_count = 60;
        static ConsoleKeyInfo anwer;
        static string FolderCK11;
        static void Main(string[] args)
        {
            Console.BackgroundColor = ConsoleColor.Black; //Console.Clear();
            hr_count = Console.WindowWidth - 1;
            string title = "Автоматизированная установка клиента ОИК СК-11";
            Console.ForegroundColor = ConsoleColor.Cyan;
            PrintHR(); Console.WriteLine(title); PrintHR();
#if DEBUG
            Console.ForegroundColor = ConsoleColor.Magenta; Console.WriteLine("Debug mode");
#endif
            Console.ForegroundColor = ConsoleColor.White;
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            Version version = assembly.GetName().Version;
            title += " ver " + (FileVersionInfo.GetVersionInfo((Assembly.GetExecutingAssembly()).Location)).ProductVersion + ": ";
            Console.Title = title;
            //string ScriptFullPathName = Application.ExecutablePath;
            //String ScriptFolder = Path.GetDirectoryName(ScriptFullPathName);                        
            Hostname = Environment.MachineName;//Hostname = Environment.GetEnvironmentVariable("COMPUTERNAME");

            #region Check CK-11 is installed
            //if (IsCK11Installed())
            //{
            //    PrintWarn("Клиент СК-11 уже установлен на этом ПК");
            //}            
            #endregion 
            #region Read setting
            ReadSetting("Distrib_Folder", ref Distrib_Folder);
            ReadSetting("Distrib_Folder_Runtime", ref Distrib_Folder_Runtime);
            ReadSetting("Distrib_Folder_CK11", ref Distrib_Folder_CK11);
            ReadSetting("autoinstaller_CK11", ref autoinstaller_CK11);
            ReadSetting("installer_CK11", ref installer_CK11);
            ReadSetting("Service_User", ref Service_User);
            ReadSetting("Service_domain", ref Service_domain);
            ReadSetting("MailServer", ref MailServer);
            ReadSetting("emailOikAdmin", ref emailOikAdmin);
            #endregion

            #region Check OS            
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("Имя компьютера {0} ", Hostname);
            try
            {
                System.DirectoryServices.ActiveDirectory.Domain cdomain = System.DirectoryServices.ActiveDirectory.Domain.GetComputerDomain();
                string PCDomain = cdomain.Name.ToUpper();
                Console.Write("Компьютер {0} в домене {1}...", Hostname, PCDomain); PrintOK();
                if (PCDomain.Contains("."))
                {
                    PCDomain = PCDomain.Substring(0, PCDomain.IndexOf('.'));
                }
                Service_domain = Service_domain.ToUpper();
                if (String.Compare(PCDomain, Service_domain, false) != 0)
                {
                    PrintWarn(String.Format("Домен компьютера {0} и домен {1} сервисного пользователя ОИК {2} не совпадают", PCDomain, Service_domain, Service_User));
                }
            }
            catch
            {
                PrintWarn(String.Format("Компьютер {0} не в домене", Hostname));
#if DEBUG
                Service_domain = Hostname;
#else
                Service_domain = String.Empty;
#endif
            }
            Console.ForegroundColor = ConsoleColor.White;
            OSInfo OSrequire = new OSInfo("Windows", "64", "10", "1049");
            OSInfo OS = new OSInfo();
            Console.Write(OS.Info() + " ");
            if (OS.Compare(OSrequire)) PrintOK(); else PrintWarn("ОС не соответсвует требованиям");
            #endregion
            #region Check Admin right and restart process
            //TODO:Проверка что текущий пользователь член группу локлаьных админов
            String ProcessOwner = String.Empty;
            Console.ForegroundColor = ConsoleColor.White;
            if (GetProcessOwner(ref ProcessOwner)) { Console.Write("Программа запущена от имени {0}. ", ProcessOwner); }
            if (IsAdministrator()) { Console.Write("Программа запущена от имени Администратора. "); PrintOK(); }
            else
            {
                PrintWarn("Программа запущена НЕ от имени Администратора");
                ProcessStartInfo proc = new ProcessStartInfo();
                proc.UseShellExecute = true;
                proc.WorkingDirectory = Environment.CurrentDirectory;
                proc.FileName = Application.ExecutablePath;
                proc.Verb = "runas";
                Console.Write("Попытка перезапустить программу от имени Администратора...");
                try
                {
#if !DEBUG 
                    Process.Start(proc);
#endif 
                    PrintOK();
#if !DEBUG 
                    IntPtr hwind = GetConsoleWindow();
                    if (hwind != IntPtr.Zero)
                    {
                        ShowWindow(hwind, 0);
                    }
                    ScriptFinish(true);
                    return;
#endif 
                }
                catch
                {
                    PrintFail();
                }

            }

            #endregion

            #region Add Service User
            Console.ForegroundColor = ConsoleColor.DarkGray; PrintHR();
            #region Get name of builtin admin group

            if (!WMIGetLocalAdminGroup(ref LocalAdministratorsGroup))
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("Не удалось определить встроенную группа локаьных администраторов");
                Console.ForegroundColor = ConsoleColor.DarkGray;
                Console.WriteLine(__Error);
                ScriptFinish(true);
            }
#if DEBUG
            Console.ForegroundColor = ConsoleColor.Gray; Console.WriteLine("Local Administrators grouop is {0}", LocalAdministratorsGroup); Console.ResetColor();
#endif
            #endregion
            #region Check service user in builtin admin group
            if (IsUserInLocalGrooup(ref Service_User, ref LocalAdministratorsGroup, Service_domain))
            {
                Console.ForegroundColor = ConsoleColor.Green; Console.WriteLine("Сервисный пользователь {0}\\{1} уже есть в локальной группе {2}",
                    Service_domain, Service_User, LocalAdministratorsGroup);
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.White; Console.Write("Добавление сервисного пользователя {0}\\{1} в локальную группу {2} ...", Service_domain.ToUpper(), Service_User.ToUpper(), LocalAdministratorsGroup);
                if (AddUserToLocalGrooup(Service_User, LocalAdministratorsGroup, Service_domain))
                    PrintOK();
                else
                {
                    PrintFail();
                    Console.ForegroundColor = ConsoleColor.DarkGray;
                    Console.WriteLine("{0}", __Error);
                }
            }
            #endregion
            #endregion
            #region Set pagefile
            bool PageFile = false;
            Console.ResetColor(); Console.WriteLine(new string(hr_char, hr_count));
            Console.ForegroundColor = ConsoleColor.White;
            ulong PhysicalMemory = GetPhysicalMemory();
            Console.Write("В компьютере установлено памяти {0} Мб ", PhysicalMemory);
            if (minPhysicalMemory <= PhysicalMemory) PrintOK();
            else
            {
                PrintWarn(String.Format("Не соответствует требованиям в {0} Мб !", minPhysicalMemory));
                Console.ResetColor(); Console.WriteLine("Продолжать установку? (Y/N) Да");
                PhysicalMemory = minPhysicalMemory;
            }
            ulong PageFileMaximumSize = (ulong)(1.5 * PhysicalMemory);
            ulong PageFileInitialSize = PageFileMaximumSize;
            ManagementObjectSearcher
            searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PageFileSetting");
            foreach (ManagementObject obj in searcher.Get())
            {
#if DEBUG
                Console.Write("{0}\t{1}\t{2}", obj["Name"].ToString(), obj["InitialSize"], obj["MaximumSize"]);
#endif                
                if (Convert.ToUInt64(obj["InitialSize"]) == PageFileInitialSize && Convert.ToUInt64(obj["MaximumSize"]) == PageFileMaximumSize)
                {
#if !DEBUG
                    Console.Write("{0}\t{1}\t{2}", obj["Name"].ToString(), obj["InitialSize"], obj["MaximumSize"]);
#endif
                    PageFile = true; PrintOK(); break;
                }
                Console.WriteLine();
            }
            if (!PageFile)
            {
                Console.Write("Установка размеров файла подкачки InitialSize={0} MaximumSize={1}...", PageFileInitialSize, PageFileMaximumSize);
                if (SetPageFileSize(PageFileInitialSize, PageFileMaximumSize)) PrintOK(); else { PrintFail(); Console.ForegroundColor = ConsoleColor.DarkGray; Console.WriteLine(__Error); }
            }
            #endregion
            #region Check Distrib Folder
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("Проверка папки дистрибутивов {0} ...", Distrib_Folder);
#if !DEBUG
            if (!Directory.Exists(Distrib_Folder)) { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine("СБОЙ"); ScriptFinish(true); }
#endif
            PrintOK();

            #endregion
            #region Check .NET
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("Проверка версии .NET ");
            String _NET_VERSION = String.Empty;
            try
            {
                _NET_VERSION = (string)Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full", "Version", String.Empty);
            }
            catch (Exception e)
            {
#if DEBUG
                __Error = e.ToString();
#else
                __Error = e.Message;
#endif
            }
            if (!String.IsNullOrEmpty(_NET_VERSION))
            {
                Console.Write(_NET_VERSION);
                if (String.Compare(_NET_VERSION, "4.8.") >= 0) PrintOK();
                else _NET_VERSION = String.Empty;
            }
            if (String.IsNullOrEmpty(_NET_VERSION))
            {
                Console.ForegroundColor = ConsoleColor.White;
                Console.Write("Установка .NET 4.8 ");
                if (RunExe(Distrib_Folder + "\\" + Distrib_Folder_Runtime + "\\" + @"ndp48-x86-x64-allos-enu.exe", "") == 0)
                {
                    PrintOK();
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.Write("\n\n\n\aНеобходимо перегрузить компьютер! Перегружать (Y/N)");
                    anwer = Console.ReadKey();
                    if (anwer.KeyChar == 121 || anwer.KeyChar == 89)
                    {
                        Console.ForegroundColor = ConsoleColor.White;
                        Console.WriteLine("\nПерезагрузка ПК ...");
                        try
                        {
                            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce", Path.GetFileName(Application.ExecutablePath), Application.ExecutablePath, RegistryValueKind.String);
                            Console.ForegroundColor = ConsoleColor.Magenta;
                            PrintWarn("\nУстановка продолжиться АВТОМАТИЧЕСКИ после перезагрузки ПК");
                        }
                        catch { }
                        System.Diagnostics.Process.Start("shutdown.exe", "-r -t 10");
                    }
                }
                else { PrintFail(); Console.WriteLine(__Error); }
            }


            #endregion

            #region Install Runtime            
            Console.ForegroundColor = ConsoleColor.DarkGray; PrintHR(); Console.ForegroundColor = ConsoleColor.White;

            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("Установка VC2010 ");
            if (RunExe(Distrib_Folder + "\\" + Distrib_Folder_Runtime + "\\" + @"VC2010_redist_x64.exe", "/install /passive /norestart") == 0)
                PrintOK();
            else { PrintFail(); Console.WriteLine(__Error); }
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("Установка VC2015-2019 ");
#if DEBUG
            PrintOK();
#else
            if (RunExe(Distrib_Folder + "\\" + Distrib_Folder_Runtime + "\\" + @"VC2015-2019_redist.x64.exe", "/install /passive /norestart") == 0)
                PrintOK();
            else { PrintFail(); Console.WriteLine(__Error); }
#endif
            #endregion


            #region Copy distrub
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("Копирование дистрибутива ");
            string FolderTMP = String.Empty;
            bool useFolderTMP = false;
            try
            {
                FolderTMP = Path.Combine(Path.GetTempPath(), Distrib_Folder_CK11);
                Directory.CreateDirectory(FolderTMP);
                Console.Write(" во временную папку " + FolderTMP);
                useFolderTMP = true;
            }
            catch (Exception e)
            {
#if DEBUG
                __Error = e.ToString();
#else
                __Error = e.Message;
#endif
                PrintFail();
                Console.WriteLine("Не удалось создать временную папку для копирования дистрибутива\n" + __Error);
            }
            if (useFolderTMP)
                if (DirectoryCopy(Distrib_Folder + "\\" + Distrib_Folder_CK11, FolderTMP, true))
                {
                    PrintOK();
                }
                else
                {
                    PrintFail();
                    Console.WriteLine(__Error);
                    PrintWarn("Не удалось скопировать дистрибутив во временную папку.\n Установка будет запущена с " + Distrib_Folder + "\\" + Distrib_Folder_CK11);
                    useFolderTMP = false;
                }
            #endregion
            #region Run INSTALL OIK
            Console.ForegroundColor = ConsoleColor.White;

            autoinstaller_CK11 = Distrib_Folder + @"\autoinstall\" + autoinstaller_CK11;
            Console.Write("Запуск программы установки ОИК СК-11 " + autoinstaller_CK11);
            string Param_autoinstaller_CK11 = "\"" + (useFolderTMP ? FolderTMP : Distrib_Folder + "\\" + Distrib_Folder_CK11) + "\" \"" + installer_CK11 + "\"";
#if DEBUG
            Console.Write("\n{0} ", Param_autoinstaller_CK11);
#endif
            if (RunExe(autoinstaller_CK11, Param_autoinstaller_CK11) == 0)
                PrintOK();
            else { PrintFail(); Console.WriteLine(__Error); }
            if (Directory.Exists(FolderTMP))
            {
                Console.ForegroundColor = ConsoleColor.White;
                Console.Write("Удаление временной папки");
                try
                {
                    Directory.Delete(FolderTMP, true);
                }
                catch { }
                if (Directory.Exists(FolderTMP)) PrintFail(); else PrintOK();
            }
            #endregion
            #region MemHeapSize    

            bool setMemHeapSize = false;
            int MemHeapSize = 4096;
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("Установка MemHeapSize {0} ", MemHeapSize);            
            string File_CK11_JSON = Path.Combine(Environment.GetEnvironmentVariable("ALLUSERSPROFILE") + @"\Monitel\CK-11", "PlatformSettings.json");
            int Timeout = 10;
            while (!(Timeout > 0 || !File.Exists(File_CK11_JSON)))
            {
                Timeout--;
                Thread.Sleep(1000);
                Console.Write(".");
            }

            if (!File.Exists(File_CK11_JSON))
            {
                PrintWarn("Файл \"" + File_CK11_JSON + "\"не найден");
            }
            else // JSON read/write interface https://www.newtonsoft.com/json/help/html/ReadingWritingJSON.htm
            {
				1.Wait for text in logfile
				
                if (SetJSONParameterValue(File_CK11_JSON, "malMemHeapSize", MemHeapSize.ToString())) PrintOK();
                else PrintWarn("Не удалось установить MemHeapSize в файл \"" + File_CK11_JSON + "\"\n" + __Error);
            }
            #endregion
            #region Send email
            //Console.Write("\a");
            //Console.ForegroundColor = ConsoleColor.Yellow;
            //Console.Write("\n\n\n\aОтправить администратору ОИК сообщение об установленном клиенте на этом ПК (Y/N)");
            //anwer = Console.ReadKey();
            //if (anwer.KeyChar == 121 || anwer.KeyChar == 89)
            //{
            //    Console.ForegroundColor = ConsoleColor.White;
            //    Console.Write("\nОтправка сообщения");
            //    if (SendMail()) PrintOK();
            //    else { PrintFail(); Console.WriteLine(__Error); }
            //}
            #endregion



            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write("\n\n\n\aРабота скрипта завершена\n\n\n");
            ScriptFinish(true);
        }

        //********************************************************
        static public int RunExe(string pProcessPath, string args)
        {
            int RunExe = -1;
            ProcessStartInfo ProcessInfo = new ProcessStartInfo();
            ProcessInfo.Arguments = args;
            //ProcessInfo.WorkingDirectory = PathDB + "\\distrib";
            ProcessInfo.FileName = pProcessPath;
            Process Process;
            string[] cProgress = { "▄█", "█▄", "█▀", "▀█" };//Console.Write("|-\\/▄█▌▐█▌░▒▓█■▬▀▄");
            int cIndex = 0;
            int cCount = cProgress.Length;
            int cSeconds = 0;
            string cMsg = "";
            try
            {
                Process = Process.Start(ProcessInfo);
                Console.CursorVisible = false;
                while (!Process.WaitForExit(1000))
                {
                    Console.Write(new string('\b', cMsg.Length));
                    cMsg = String.Format("{0} {1} сек", cProgress[cIndex++], cSeconds++);
                    Console.Write(cMsg);
                    if (cIndex == cCount) cIndex = 0;
                    //Console.Write(new string('\b', cMsg.Length));
                }

#if DEBUG
                Console.Write("\nExitCode={0} ", Process.ExitCode);
#endif
                RunExe = Process.ExitCode;
            }
            catch (Exception e)
            {
#if DEBUG
                __Error = e.ToString();
#else
                __Error = e.Message;
#endif
            }
            Console.CursorVisible = true;
            return RunExe;
        }
        public static bool DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            bool boolDirectoryCopy = false;
            try
            {
                DirectoryInfo dir = new DirectoryInfo(sourceDirName);
                if (!dir.Exists)
                {
                    __Error = "Исходная папка не найдена";
                }
                else
                {
                    DirectoryInfo[] dirs = dir.GetDirectories();
                    // If the destination directory doesn't exist, create it.       
                    Directory.CreateDirectory(destDirName);
                    // Get the files in the directory and copy them to the new location.
                    FileInfo[] files = dir.GetFiles();
                    foreach (FileInfo file in files)
                    {
                        string tempPath = Path.Combine(destDirName, file.Name);
                        file.CopyTo(tempPath, false);
                    }
                    // If copying subdirectories, copy them and their contents to new location.
                    boolDirectoryCopy = true;
                    if (copySubDirs)
                    {
                        foreach (DirectoryInfo subdir in dirs)
                        {
                            string tempPath = Path.Combine(destDirName, subdir.Name);
                            if (!DirectoryCopy(subdir.FullName, tempPath, copySubDirs))
                            { boolDirectoryCopy = false; break; }
                        }
                    }

                }
            }
            catch (Exception e)
            {
#if DEBUG
                __Error = e.ToString();
#else
                __Error = e.Message;
#endif
            }
            return boolDirectoryCopy;
        }
        //****************************************************
        public static void ScriptFinish(bool pause)
        //****************************************************			
        {

            if (pause)
            {
                Console.ForegroundColor = ConsoleColor.White; Console.WriteLine(); Console.Write("Press any key to exit . . . "); Console.ResetColor();
#if DEBUG
                int timeout = 100;
#else
int timeout =10;
#endif
                DateTime timeoutvalue = DateTime.Now.AddSeconds(timeout);
                while (DateTime.Now < timeoutvalue)
                {
                    if (Console.KeyAvailable) break;
                    Thread.Sleep(100);
                }
            }
            Environment.Exit(1);
        }
        #region GetLocalAdminGroup ver 1
        //*******************************************
        static public bool WMIGetLocalAdminGroup(ref string group)
        //*******************************************
        {
            bool GetLocalAdminGroup = false;
            try
            {
                ManagementObjectSearcher searcher =
                    new ManagementObjectSearcher(__root_CIMv2, "SELECT Name FROM Win32_Group WHERE SID='S-1-5-32-544'");
                ManagementObjectCollection Collection = searcher.Get();
                //if (Collection.Count != 0)
                {
                    foreach (ManagementObject Obj in Collection)
                    {
                        group = Obj["Name"].ToString();
                        GetLocalAdminGroup = true;
                        break;
                    }
                }
            }
            catch (ManagementException e)
            {
                __Error = e.ToString();
            }
            return (GetLocalAdminGroup);
        }
        #endregion
        #region GetLocalAdminGroup ver 2
        static public bool GetLocalAdminGroup2(ref string group)
        {
            bool GetLocalAdminGroup = false;
            using (DirectoryEntry d = new DirectoryEntry("WinNT://" + Hostname + "/SID=S-1-5-32-544")) //https://www.script-coding.com/ADSI.html
            {
                {
                    object members = d.Invoke("Members", null);
                    foreach (object member in (IEnumerable)members)
                    {
                        DirectoryEntry x = new DirectoryEntry(member);
                        Console.Out.WriteLine(x.Name);
                    }
                }
                GetLocalAdminGroup = true;
            }

            return (GetLocalAdminGroup);
        }
        #endregion
        #region GetLocalAdminGroup ver 3
        static public bool GetLocalAdminGroup3(ref string group)
        {
            bool GetLocalAdminGroup = false;
            try
            {
                group = new SecurityIdentifier("S-1-5-32-544").Translate(typeof(NTAccount)).Value.Split('\\')[1];
                GetLocalAdminGroup = true;
            }//https://docs.microsoft.com/en-us/dotnet/api/system.security.principal.securityidentifier.translate?view=net-5.0
            //catch (ArgumentNullException e) { __Error = "ArgumentNullException\n"+ e.ToString(); }
            //catch (Exception e) when (e is ArgumentException) { __Error = "ArgumentException\n" + e.ToString(); }
            //catch (IdentityNotMappedException e) { __Error = "IdentityNotMappedException\n"+ e.ToString(); }
            //catch (Exception e) when (e is SystemException) { __Error = "SystemException\n" + e.ToString(); }
            catch (Exception e) { __Error = "GetLocalAdminGroup()" + e.ToString(); }
            return (GetLocalAdminGroup);
        }
        #endregion
        //****************************************************
        static public bool IsUserInLocalGrooup(ref String User, ref String Group, String Domain)
        {
            bool IsDomainUserInLocalGrooup = false;
            //#if DEBUG
            //Console.ForegroundColor = ConsoleColor.Gray;
            //Console.WriteLine("Group {0} contain:", LocalAdministratorsGroup);
            //#endif
            try
            {
                StringBuilder sBuilder = new StringBuilder("GroupComponent=");
                sBuilder.Append('"');
                sBuilder.Append("Win32_Group.Domain=");
                sBuilder.Append("'");
                sBuilder.Append(Hostname);
                sBuilder.Append("'");
                sBuilder.Append(",Name=");
                sBuilder.Append("'");
                sBuilder.Append(LocalAdministratorsGroup);
                sBuilder.Append("'");
                sBuilder.Append('"');
                String[] properties = { "PartComponent" };
                SelectQuery sQuery = new SelectQuery("Win32_GroupUser", sBuilder.ToString(), properties);
                ManagementObjectSearcher Searcher = new ManagementObjectSearcher(sQuery);
                //ManagementObjectSearcher searcher = new ManagementObjectSearcher(__root_CIMv2,
                //"SELECT * FROM Win32_GroupUser WHERE Win32_Group.Domain = 'SKORIK10', Name = 'Администраторы'");
                //"SELECT PartComponent  FROM Win32_GroupUser WHERE GroupComponent=\"Win32_Group.Domain = 'SKORIK10', Name = 'Администраторы'\"");
                // проверенный SELECT PartComponent  FROM Win32_GroupUser WHERE GroupComponent="Win32_Group.Domain='SKORIK10',Name='Администраторы'"                
                ManagementObjectCollection Collection = Searcher.Get();
                foreach (ManagementObject Obj in Collection)
                {

                    ManagementPath path = new ManagementPath(Obj["PartComponent"].ToString());
                    if (path.ClassName == "Win32_UserAccount")
                    {
                        //Console.WriteLine(path.ToString());
                        //Console.WriteLine(path.RelativePath);
                        String[] fields = path.RelativePath.Split(',');
                        //Console.WriteLine("{0} === {1}",fields[0], fields[1]);
                        String domain = fields[0].Substring(fields[0].IndexOf("=") + 1).Replace('"', ' ').Trim();
                        String user = fields[1].Substring(fields[1].IndexOf("=") + 1).Replace('"', ' ').Trim();
                        //#if DEBUG
                        //Console.WriteLine("{0}\\{1}", domain, user);
                        //#endif

                        if (String.Compare(Service_domain, domain, true) == 0) if (String.Compare(Service_User, user, true) == 0)
                            {
                                IsDomainUserInLocalGrooup = true;
                                break;
                            }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            return (IsDomainUserInLocalGrooup);
        }
        static public bool AddUserToLocalGrooup(string User, string Group, string domain)
        {
#if DEBUG
            if (!IsAdministrator()) return false;
#endif
            bool AddUserToLocalGrooup = false;
            try
            {
                PrincipalContext M = new PrincipalContext(ContextType.Machine);
                if (M == null) { __Error = "Не удалось подключиться к локальному списку пользователей"; }
                else
                {
                    GroupPrincipal G = GroupPrincipal.FindByIdentity(M, IdentityType.Sid, "BA");//https://docs.microsoft.com/en-us/windows/win32/secauthz/sid-strings
                    if (G == null) { __Error = "Не удалось подключиться к группе локальных администраторов"; }
                    else
                    {
                        PrincipalContext D = new PrincipalContext(ContextType.Domain, domain);
                        if (D == null) { __Error = "Не удалось подключиться к домену"; }
                        else
                        {
                            //G.Members.Add(D, IdentityType.SamAccountName, "User123");
                            //G.Members.Add(D, IdentityType.Name, domain + "\\" + User);
                            G.Members.Add(D, IdentityType.Name, User);
                            G.Save();
                            G.Dispose();
                            AddUserToLocalGrooup = IsUserInLocalGrooup(ref User, ref Group, domain);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
#if DEBUG
                __Error = ex.ToString();
#else
                __Error = ex.Message;
#endif
            }
            return AddUserToLocalGrooup;
        }



        static public ulong GetPhysicalMemory()
        {
            ulong GetPhysicalMemory = 0;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Capacity  FROM Win32_PhysicalMemory");
            foreach (ManagementObject obj in searcher.Get())//Console.WriteLine(obj["Capacity"].ToString());
                GetPhysicalMemory += Convert.ToUInt64(obj["Capacity"].ToString()) >> 20;
            return GetPhysicalMemory;
        }
        static void PrintOK() { Console.BackgroundColor = ConsoleColor.Green; Console.ForegroundColor = ConsoleColor.White; Console.Write(" OK "); Console.ResetColor(); Console.WriteLine("."); }
        static void PrintFail() { Console.BackgroundColor = ConsoleColor.Red; Console.ForegroundColor = ConsoleColor.White; Console.Write(" СБОЙ "); Console.ResetColor(); Console.WriteLine("."); }
        static void PrintWarn(string Msg) { Console.BackgroundColor = ConsoleColor.Yellow; Console.ForegroundColor = ConsoleColor.Black; Console.Write(Msg); Console.ResetColor(); Console.WriteLine("."); }
        static void PrintHR() { Console.WriteLine(new string(hr_char, hr_count)); }
        static public bool SetPageFileSize(ulong PageFileInitialSize, ulong PageFileMaximumSize)
        {
            bool SetPageFileSize = false;
#if !DEBUG
            try
            {
                RegistryKey rk = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management", true);
                rk.SetValue("PagingFiles", new string[] { String.Format("{0}\\pagefile.sys {1} {2}", "c:", PageFileInitialSize, PageFileMaximumSize) },
                    RegistryValueKind.MultiString);
                SetPageFileSize = true;
            }
            catch (Exception e) { __Error = e.Message; }
#endif
            return SetPageFileSize;
        }

        #region Principal Function
        //https://wiki.plecko.hr/doku.php?id=windows:ad:ad.net
        //private string sDomain = "test.com";
        //private string sDefaultOU = "OU=Test Users,OU=Test,DC=test,DC=com";
        //private string sDefaultRootOU = "DC=test,DC=com";
        //private string sServiceUser = @"ServiceUser";
        //private string sServicePassword = "ServicePassword";
        static public bool AddUserToGroup(string sUserName, string sGroupName)
        {
            bool AddUserToGroup = false;
            try
            {
                UserPrincipal oUserPrincipal = GetUser(sUserName);
                if (oUserPrincipal == null) __Error = String.Format("User {0} not found", sUserName);
                else
                {

                    PrincipalContext oPrincipalContext = new PrincipalContext(ContextType.Domain);//, Hostname);
                    GroupPrincipal oGroupPrincipal = GroupPrincipal.FindByIdentity(oPrincipalContext, sGroupName);




                    if (oGroupPrincipal == null) __Error = String.Format("Group {0} not found", sGroupName);
                    else
                    {
                        if (!IsUserGroupMember(sUserName, sGroupName))
                        {
                            oGroupPrincipal.Members.Add(oUserPrincipal);
                            oGroupPrincipal.Save();
                            AddUserToGroup = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                __Error = ex.Message;
            }
            return AddUserToGroup;
        }
        static public bool IsUserGroupMember(string sUserName, string sGroupName)
        {
            UserPrincipal oUserPrincipal = GetUser(sUserName);
            GroupPrincipal oGroupPrincipal = GetGroup(sGroupName);

            if (oUserPrincipal != null && oGroupPrincipal != null)
            {
                return oGroupPrincipal.Members.Contains(oUserPrincipal);
            }
            else
            {
                return false;
            }
        }
        static public GroupPrincipal GetGroup(string sGroupName)
        {
            PrincipalContext oPrincipalContext = GetPrincipalContext();

            GroupPrincipal oGroupPrincipal = GroupPrincipal.FindByIdentity(oPrincipalContext, sGroupName);
            return oGroupPrincipal;
        }
        static public UserPrincipal GetUser(string sUserName)
        {
            PrincipalContext oPrincipalContext = GetPrincipalContext();

            UserPrincipal oUserPrincipal = UserPrincipal.FindByIdentity(oPrincipalContext, sUserName);
            return oUserPrincipal;
        }
        static public PrincipalContext GetPrincipalContext()
        {
            //PrincipalContext oPrincipalContext = new PrincipalContext(ContextType.Domain, sDomain, sDefaultOU, ContextOptions.SimpleBind, sServiceUser, sServicePassword);
            PrincipalContext oPrincipalContext = new PrincipalContext(ContextType.Domain, Service_domain);
            return oPrincipalContext;
        }
        #endregion
        public static bool IsAdministrator()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
        static public bool GetProcessOwner(ref string owner)
        {
            bool GetProcessOwner = false;
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT  * FROM Win32_Process WHERE Handle=\"" + System.Diagnostics.Process.GetCurrentProcess().Id.ToString() + "\"");
                foreach (ManagementObject process in searcher.Get())
                {
                    string[] argList = new string[] { string.Empty, string.Empty };
                    int returnVal = Convert.ToInt32(process.InvokeMethod("GetOwner", argList));
                    if (returnVal == 0)
                    {
                        owner = argList[1] + "\\" + argList[0];
                        GetProcessOwner = true;
                    }
                    else { __Error = String.Format("Owner not found"); }
                }
            }
            catch (Exception ex)
            {
                __Error = ex.ToString();
            }
            return GetProcessOwner;
        }
        static void ReadAllSettings()
        {
            try
            {
                var appSettings = ConfigurationManager.AppSettings;

                if (appSettings.Count == 0)
                {
                    Console.WriteLine("AppSettings is empty.");
                }
                else
                {
                    foreach (var key in appSettings.AllKeys)
                    {
                        Console.WriteLine("Key: {0} Value: {1}", key, appSettings[key]);
                    }
                }
            }
            catch (ConfigurationErrorsException)
            {
                Console.WriteLine("Error reading app settings");
            }
        }
        static void ReadSetting(string key, ref string variable)
        {
            try
            {
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.Write(key); Console.Write(" ");
                var appSettings = ConfigurationManager.AppSettings;
                if (appSettings[key] == null) throw new Exception();
                variable = appSettings[key];
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine(variable);
            }
            catch /*(ConfigurationErrorsException)*/
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(variable);
            }
        }
        public static bool SendMail()
        {
            bool SendMail = false;
            string server = "mail";
            string to = "skorikoff@yug.so-ups.ru";
            string from = "ben@sof.com";
            MailMessage message = new MailMessage(from, to);
            message.Subject = "Using the new SMTP client.";
            message.Body = @"Using this new feature, you can send an email message from an application very easily.";
            SmtpClient client = new SmtpClient(server);
            // Credentials are necessary if the server requires the client
            // to authenticate before it will send email on the client's behalf.
            client.UseDefaultCredentials = true;
            try
            {
                client.Send(message);
                SendMail = true;
            }
            catch (Exception e)
            {
#if DEBUG
                __Error = e.ToString();
#else
                __Error = e.Message;
#endif
            }
            return SendMail;
        }
        public static bool IsCK11Installed()
        {
            bool IsCK11Installed = false;//SELECT* FROM Win32_InstalledWin32Program WHERE(Name LIKE  '%СК-11%' AND Vendor LIKE '%Монитор Электрик%')
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(
                "SELECT * FROM Win32_InstalledWin32Program WHERE (Name LIKE  '%СК-11%' AND Vendor LIKE '%Монитор Электрик%')");
                //"SELECT * FROM Win32_InstalledWin32Program WHERE(Name LIKE  \"%СК-11%\" AND Vendor LIKE \"%Монитор Электрик%\"");
                //"SELECT  * FROM Win32_Process WHERE Handle=\"" + System.Diagnostics.Process.GetCurrentProcess().Id.ToString() + "\"");
                foreach (ManagementObject app in searcher.Get())
                {

                    //Console.WriteLine("Name = {0} Vendor = {1]",app["Name"].ToString(), app["Vendor"].ToString());
                    IsCK11Installed = true;
                    break;
                }
            }
            catch (Exception e)
            {
#if DEBUG
                __Error = e.ToString();
#else
                __Error = e.Message;
#endif
            }
            return IsCK11Installed;
        }
        //************************************
        public static bool SetJSONParameterValue(String File, String ParameterName, String newParameterValue)        {
            // 0  ...   S-1 | S    .......          |,| ..............
            //               _______ L _________ 
            //
            //              S .... |:|  old_value
            //              ___v ___
            //
            // 0  ...    S-1| S    |:| new_value    |,| S+L+1 .......
            bool result = false;
            try
            {
                StreamReader sr = new StreamReader(File);                
                string input = sr.ReadToEnd();
                sr.Close();
                string pattern = @""" *" + ParameterName + @" *"" *: *. *[,}]";
                Match m1 = Regex.Match(input, pattern);
                if (m1.Success)
                {
                    pattern = @""" *" + ParameterName + @" *"" *:";
                    Match m2 = Regex.Match(m1.Value, pattern);
                    if (m2.Success)
                    {
                        string newValue =
                            input.Substring(0, m1.Index + m2.Length)
                            +
                            " " + newParameterValue
                            +
                            input.Substring(m1.Index + m1.Length - 1);
                        StreamWriter sw = new StreamWriter(File, false);
                        sw.Write(newValue);
                        sw.Close();
                        result = true;
                    }
                    else  { __Error = "Значение поля " + ParameterName + "не найдено"; }
                }
                else { __Error = "Поле " + ParameterName + "не найдено"; }
            }
            catch (Exception e)
            {
#if DEBUG
                __Error = e.ToString();
#else
                __Error = e.Message;
#endif
            }
            return result;
        }
    }
}



/*
1
 Справочные данный
WinNT ADsPath   https://docs.microsoft.com/en-us/windows/win32/adsi/winnt-adspath 
Пользователь  - админ  https://www.meziantou.net/check-if-the-current-user-is-an-administrator.htm#check-if-the-current 
Процесс  - админ  https://www.meziantou.net/check-if-the-current-user-is-an-administrator.htm#check-if-the-current 
Run a process as Administrator  https://daoudisamir.com/run-a-process-as-administrator-with-c-programmatically/
Повышение привилегий процесса программно - https://generacodice.com/ru/articolo/39156/In-Perforce%2C-can-you-rename-a-folder-to-the-same-name-but-cased-differently
IsInRole(WindowsBuiltInRole.Administrator - http://engram404.net/how-to-request-uac-user-account-control-elevated-permissions/
 
 */
