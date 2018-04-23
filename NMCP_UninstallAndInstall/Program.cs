using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using System.Threading;

namespace NMCP_UninstallAndInstall
{
    class Program
    {
        static void Main(string[] args)
        {
            Automation.Uninstall();
            Automation.DeleteLatestFolder();
            Automation.CopyLatestBuildWithBat();
            Automation.Install();
        }
    }

    public class Automation
    {
        public static string OriginalFolder = @"\\SHAL0560\share\builds\Machinery\setup\Client";
        public static string DestinationFolder = @"C:\Perforce\NMCP\Nauticus\Products\main\Machinery\GuiAutomationTest\Setup";

        public static string GetLatestBuildFolder()
        {
            string latestBuildFolder = string.Empty;
            DirectoryInfo dInfo = new DirectoryInfo(OriginalFolder);
            DirectoryInfo[] dirs = dInfo.GetDirectories();

            DirectoryInfo temp = dirs[0];
            foreach (DirectoryInfo dir in dirs)
            {
                if (dir.CreationTime >= temp.CreationTime)
                {
                    temp = dir;
                }
            }

            latestBuildFolder = temp.Name;
            return latestBuildFolder;
        }

        public static void CopyLatestBuildWithBat()
        {
            Console.WriteLine("Start to copy build");
            var commandLine = "xcopy ";
            commandLine = commandLine + "\"" + OriginalFolder + "\"";
            commandLine = commandLine + " " + DestinationFolder;
            commandLine = commandLine + @" /E";

            string str = System.Environment.CurrentDirectory;
            str = str + @"\CopyBuild.bat";

            StreamWriter sw = new StreamWriter(str, false);
            sw.WriteLine(commandLine);
            sw.Close();

            Process p = new System.Diagnostics.Process();
            p.StartInfo.UseShellExecute = false;
            //p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.FileName = str;
            p.Start();
            //p.WaitForExit();
            //ExecuteWithoutWait(commandLine);
            //Console.WriteLine("Pop up the copying process");
            Thread.Sleep(180000);
            Console.WriteLine("Copy the latest build to the Setup folder correctly");
        }

        public static void DeleteLatestFolder()
        {
            Console.WriteLine("Start to delete the 'Latest' folder");
            DirectoryInfo dInfo = new DirectoryInfo(DestinationFolder);
            //dInfo.Delete();

            // delete files            
            DeleteFiles(dInfo.FullName);

            // delete sub folders
            foreach (var dir in dInfo.GetDirectories())
            {
                DeleteFiles(dir.FullName);
                foreach (var subDir in dir.GetDirectories())
                {
                    DeleteFiles(subDir.FullName);
                    Directory.Delete(subDir.FullName);
                }
                Directory.Delete(dir.FullName);
            }

            // delete sub folders
            //dInfo.Delete();
            //Console.WriteLine("The 'Latest' folder has been deleted correctly");
        }

        public static void Install()
        {
            Console.WriteLine("Start to intall App");
            var commandLine = "msiexec.exe /i ";
            commandLine = commandLine + "\"" + GetMsiPath() + "\"";
            commandLine = commandLine + " /passive /l*v msiexec_log.txt";

            Execute(commandLine);
            Thread.Sleep(5000);
            Console.WriteLine("Install app completed");
        }

        public static void Uninstall()
        {
            Console.WriteLine("Start to unintall App");
            var commandLine = "msiexec.exe /x ";
            commandLine = commandLine + "\"" + GetMsiPath() + "\"";
            commandLine = commandLine + " /passive";

            Execute(commandLine);
            Thread.Sleep(1000);
            Console.WriteLine("Uninstall app completed");
        }

        public static string GetMsiPath()
        {
            string[] filedir = Directory.GetFiles(DestinationFolder, "*.msi");
            return filedir[0];
        }

        private static void ClearFolderAction()
        {
            var path = System.Environment.GetEnvironmentVariable("RanorexFolder");
            path = path + @"\SafetiOffshoreSmokeTest";

            DirectoryInfo dInfo = new DirectoryInfo(path);

            // delete files            
            DeleteFiles(dInfo.FullName);

            // delete folders
            foreach (var dir in dInfo.GetDirectories())
            {
                DeleteFiles(dir.FullName);
                foreach (var subDir in dir.GetDirectories())
                {
                    DeleteFiles(subDir.FullName);
                    Directory.Delete(subDir.FullName);
                }
                Directory.Delete(dir.FullName);
            }
        }

        private static void DeleteFiles(string path)
        {
            // delete files
            DirectoryInfo dInfo = new DirectoryInfo(path);
            foreach (var fs in dInfo.GetFiles())
            {
                var filePath = path + @"\" + fs.Name;
                File.Delete(filePath);
            }
        }

        public static void Execute(string command, int seconds = 0)
        {
            if (command != null && !command.Equals(""))
            {
                Process process = new Process();                        //创建进程对象  
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = "cmd.exe";                         //设定需要执行的命令  
                startInfo.Arguments = "/C " + command;                          //“/C”表示执行完命令后马上退出  
                startInfo.UseShellExecute = false;                      //是否使用操作系统shell启动  
                startInfo.RedirectStandardInput = false;                //不重定向输入  
                startInfo.RedirectStandardOutput = true;                //重定向输出  
                startInfo.CreateNoWindow = true;                        //不创建窗口  
                process.StartInfo = startInfo;
                try
                {
                    if (process.Start())                                //开始进程  
                    {
                        if (seconds == 0)
                        {
                            process.WaitForExit();                      //这里无限等待进程结束  
                        }
                        else
                        {
                            process.WaitForExit(seconds);               //等待进程结束，等待时间为指定的毫秒  
                        }
                    }
                }
                catch
                {

                }
                finally
                {
                    if (process != null)
                    {
                        process.Close();
                    }
                }
            }
        }

        public static void ExecuteWithoutWait(string command)
        {
            if (command != null && !command.Equals(""))
            {
                Process process = new Process();                        //创建进程对象  
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = "cmd.exe";                         //设定需要执行的命令  
                startInfo.Arguments = "/C " + command;                          //“/C”表示执行完命令后马上退出  
                startInfo.UseShellExecute = false;                      //是否使用操作系统shell启动  
                startInfo.RedirectStandardInput = false;                //不重定向输入  
                startInfo.RedirectStandardOutput = true;                //重定向输出  
                startInfo.CreateNoWindow = true;                        //不创建窗口  
                process.StartInfo = startInfo;
                process.Start();
                Thread.Sleep(900000);
            }
        }
    }
}
