﻿using System;
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
using System.Windows;
using System.Security.Principal;
using System.DirectoryServices.ActiveDirectory;
using System.DirectoryServices.AccountManagement;
using System.Data;
using System.Configuration;


namespace Install_CK11
{
    class Program
    {
        public static string __Error;
        const string __root_CIMv2 = "root\\CIMV2";
        static string Distrib_Folder = "\\\\fs2-oduyu\\CK2007\\СК-11\\";
        static string Service_User="Svc-ck11cl-oduyu";
        static string Service_domain="oduyu";
        static string Hostname;
        static string LocalAdministratorsGroup = null;
        static ulong minPhysicalMemory =6144;
        static void Main(string[] args)
        {
            Console.BackgroundColor = ConsoleColor.Black; Console.Clear();          

            const char hr_char = '─';
            int hr_count = Console.WindowWidth-1;            
            string title = "Автоматизированная установка клиента ОИК СК-11";
            Console.ForegroundColor = ConsoleColor.Cyan;            
            Console.WriteLine(new string(hr_char, hr_count));
            Console.WriteLine(title);
            Console.WriteLine(new string(hr_char, hr_count));
#if DEBUG
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine("Debug mode");
#endif
            Console.ForegroundColor = ConsoleColor.White;
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            Version version = assembly.GetName().Version;
            title += " ver " + (FileVersionInfo.GetVersionInfo((Assembly.GetExecutingAssembly()).Location)).ProductVersion + ": ";
            Console.Title = title;
            //string ScriptFullPathName = Application.ExecutablePath;
            //String ScriptFolder = Path.GetDirectoryName(ScriptFullPathName);
            //TODO:Проверка доступности сетевой папки с дитрибами
            //Hostname = Environment.GetEnvironmentVariable("COMPUTERNAME");
            Hostname = Environment.MachineName;
            //TODO:Проверка что ПК в домене 
            //Console.Write("Компьютер {0} в домене {1}...", Hostname,Service_domain);
#region Check OS            
            Console.WriteLine(GetOSFriendlyName());
#endregion
#region Check Distrib Folder
            Console.Write("Проверка папки дистрибутивов {0} ...", Distrib_Folder);
#if !DEBUG
            if (!Directory.Exists(Distrib_Folder)){Console.ForegroundColor = ConsoleColor.Red;Console.WriteLine("СБОЙ");ScriptFinish(true);}
#endif
            PrintOK();
            
#endregion
#region Add Service User
#region Get name of builtin admin group
            
            if (!WMIGetLocalAdminGroup(ref LocalAdministratorsGroup))
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("Не удалось определить встроенную группа локаьных администраторов");
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine(__Error);
                ScriptFinish(true);
            }            
#if DEBUG
            Console.ForegroundColor = ConsoleColor.Gray;Console.WriteLine("Local Administrators grouop is {0}", LocalAdministratorsGroup);Console.ResetColor();
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
                //if (AddUserToLocalGrooup(ref Service_User, ref LocalAdministratorsGroup, ref Service_domain)) 
                if (AddUserToGroup(Service_domain+ "\\"+Service_User, LocalAdministratorsGroup))
                    PrintOK();
                else { PrintFail();
                    Console.WriteLine("{0}", __Error);
                }
            }            
            #endregion
            #endregion
            #region Set pagefile
            Console.ResetColor(); Console.WriteLine(new string(hr_char, hr_count));
            ulong PhysicalMemory = GetPhysicalMemory();
            Console.Write("В компьютере установлено памяти {0} Мб ", PhysicalMemory);
            if (minPhysicalMemory <= PhysicalMemory) PrintOK();
            else
            {
                PrintWarn(String.Format("Не соответствует требованиям в {0} Мб !", minPhysicalMemory));
                Console.ResetColor(); Console.WriteLine("Продолжать установку? (Y/N) Да");
                PhysicalMemory = minPhysicalMemory;
            }
            ulong PageFileMaximumSize = (ulong) 1.5 * PhysicalMemory;
            ulong PageFileInitialSize = PageFileMaximumSize;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PageFileSetting");
            foreach (ManagementObject obj in searcher.Get())
            {
                Console.WriteLine("{0}\t{1}\t{2}",obj["Name"].ToString(), obj["InitialSize"], obj["MaximumSize"]);

                //break;
            }
            #endregion
            ScriptFinish(true);
        }

        //****************************************************
        public static void ScriptFinish(bool pause)        
        {
        //****************************************************			
            if (pause)
            {
                Console.ForegroundColor = ConsoleColor.White;Console.WriteLine();Console.Write("Press any key to exit . . . "); Console.ResetColor();
                DateTime timeoutvalue = DateTime.Now.AddSeconds(100);
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
            catch (Exception e) { __Error = "GetLocalAdminGroup()"+e.ToString(); }
            return (GetLocalAdminGroup);
        }
#endregion
        //****************************************************
        static public bool IsUserInLocalGrooup(ref String User, ref String Group, String Domain) {
        //****************************************************
            /* work ok
                using (DirectoryEntry d = new DirectoryEntry("WinNT://" + Hostname + ",computer")) //https://www.script-coding.com/ADSI.html
                {
                    using (DirectoryEntry g = d.Children.Find("Администраторы", "group"))
                    {
                        object members = g.Invoke("Members", null);
                        foreach (object member in (IEnumerable)members)
                        {
                            DirectoryEntry x = new DirectoryEntry(member);                        
                            Console.Out.WriteLine(x.Path,x.Name);
                        }
                    }
                }
                */
            bool IsDomainUserInLocalGrooup = false;
#if DEBUG
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine("Group {0} contain:", LocalAdministratorsGroup);
#endif
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
#if DEBUG
                        Console.WriteLine("{0}\\{1}", domain, user);
#endif

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
        static public bool AddUserToLocalGrooup(ref string User, ref string Group, ref string domain)
        {
        bool AddUserToLocalGrooup = false;
            try
            {
                DirectoryEntry groupEntry = new DirectoryEntry("WinNT://"+Hostname+"/" + Group + ",group");
                DirectoryEntry userEntry  = new DirectoryEntry("WinNT://" + domain + "/" + User + ",user");
                groupEntry.Children.Add(userEntry.Path.ToString(), "user");
                //groupEntry.Invoke("Add", "WinNT://"+domain + "/" + User, "user");
                //groupEntry.CommitChanges();
                AddUserToLocalGrooup =IsUserInLocalGrooup(ref User, ref Group, domain);
            }            
            catch (Exception ex)
            {
                __Error=ex.ToString();
            }
            return AddUserToLocalGrooup;
        }        
    
        public static string GetOSFriendlyName()
        {
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
            //
            string result = string.Empty;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
            foreach (ManagementObject os in searcher.Get())
            {
                result = os["Caption"].ToString();
                break;
            }
            return result;
        }
        static public ulong GetPhysicalMemory()
        {
            ulong GetPhysicalMemory = 0;            
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Capacity  FROM Win32_PhysicalMemory");
            foreach (ManagementObject obj in searcher.Get())
                //Console.WriteLine(obj["Capacity"].ToString());
                GetPhysicalMemory += Convert.ToUInt64(obj["Capacity"].ToString())>>20 ;            
            return GetPhysicalMemory;
        }
        static void PrintOK() {Console.BackgroundColor = ConsoleColor.Green;Console.WriteLine(" OK ");Console.ResetColor();}
        static void PrintFail() { Console.BackgroundColor = ConsoleColor.Red; Console.WriteLine(" СБОЙ "); Console.ResetColor(); }
        static void PrintWarn(string Msg) { Console.BackgroundColor = ConsoleColor.Yellow; Console.ForegroundColor = ConsoleColor.Black;Console.WriteLine(Msg); Console.ResetColor(); }


        //*******************************************************************************************
        //string ou = "OU=Collections,DC=Domain,DC=local";
        //PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "Domain.Local", ou);
        //static public bool AddUserToGroup(PrincipalContext ctx, DirectoryEntry userId, string groupName)
        //{
        //    try
        //    {
        //        GroupPrincipal groupPrincipal = GroupPrincipal.FindByIdentity(ctx, groupName);
        //        if (groupPrincipal != null)
        //        {
        //            DirectoryEntry entry = (DirectoryEntry)groupPrincipal.GetUnderlyingObject();
        //            entry.Invoke("Add", new object[] { userId.Path.ToString() });
        //            userId.CommitChanges();
        //       }
        //        else
        //        {
         //           return true;
        //        }
        //
        //        return true;
        //    }
        //    catch
        //    {
        //        return false;
        //    }
        //}




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
                __Error = ex.ToString();
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

    }
}









/*title = "Ну удалось определить встроенную группа локаьных администраторов\nIdentityNotMappedException\nSystem.Security.Principal.IdentityNotMappedException: Некоторые или ссылки на свойства нельзя преобразовать.\n   в System.Security.Principal.SecurityIdentifier.Translate(IdentityReferenceCollection sourceSids, Type targetType, Boolean forceSuccess)\n   в System.Security.Principal.SecurityIdentifier.Translate(Type targetType)\n   в Install_CK11.Program.GetLocalAdminGroup(String & group) в Program.cs:строка 186";
            Console.ForegroundColor = ConsoleColor.DarkCyan;
            Console.WriteLine(title);
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine(title);
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine(title);
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine(title);
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine(title);
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(title);
            Console.ForegroundColor = ConsoleColor.DarkRed;
            Console.WriteLine(title);
            Console.ForegroundColor = ConsoleColor.DarkMagenta;
            Console.WriteLine(title);
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine(title);*/



/*
 Справочные данный
WinNT ADsPath   https://docs.microsoft.com/en-us/windows/win32/adsi/winnt-adspath 
 
 
 */