using System;
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
        static void Main(string[] args)
        {
            Console.BackgroundColor = ConsoleColor.Black; Console.Clear();          

            const char hr_char = '─';
            const byte hr_count = 60;
            string title = "Автоматизированная установка клиента ОИК СК-11";
            
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


            Console.ForegroundColor = ConsoleColor.Cyan;            
            Console.WriteLine(new string(hr_char, hr_count));
            Console.WriteLine(title);
            Console.WriteLine(new string(hr_char, hr_count));
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
            #region Check Distrib Folder
            Console.Write("Проверка папки дистрибутивов {0} ...", Distrib_Folder);
            if (!Directory.Exists(Distrib_Folder))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("СБОЙ");
                ScriptFinish(true);
            }
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(" OK");
            Console.ForegroundColor = ConsoleColor.White;
            #endregion
            #region Add Service User
            #region Get name of builtin admin group
            string LocalAdministratorsGroup=null;
            if (!WMIGetLocalAdminGroup(ref LocalAdministratorsGroup))
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("Не удалось определить встроенную группа локаьных администраторов");
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine(__Error);
                ScriptFinish(true);
            }            
#if DEBUG
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine("Local Administrators grouop is {0}", LocalAdministratorsGroup);
            Console.ResetColor();
#endif
            #endregion
            #region Check service user in builtin admin group
            
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
            try
            {
#if DEBUG
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine("Group {0} contain:", LocalAdministratorsGroup);
#endif
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
                String[] properties =            {"PartComponent"};
                SelectQuery sQuery = new SelectQuery("Win32_GroupUser", sBuilder.ToString(), properties);
                #endregion
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
                        String domain=fields[0].Substring(fields[0].IndexOf("=") + 1).Replace('"', ' ').Trim();
                        String user =fields[1].Substring(fields[1].IndexOf("=") + 1).Replace('"', ' ').Trim();
#if DEBUG
                        Console.WriteLine("{0}\\{1}", domain, user);
#endif
                    
                    if (String.Compare(Service_domain, domain, true)==0) if (String.Compare(Service_User, user, true) == 0)
                            {
                        break; }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
#if DEBUG
            Console.ResetColor();
#endif


            ScriptFinish(true);
            #endregion       


#if DEBUG
#endif

            
            ScriptFinish(true);
        }

        //****************************************************
        public static void ScriptFinish(bool pause)        
        {
        //****************************************************			
            if (pause)
            {
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine();
                Console.Write("Press any key to exit . . . "); Console.ResetColor();
                DateTime timeoutvalue = DateTime.Now.AddSeconds(10);
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
    }
}
