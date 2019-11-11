using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading;
using IWshRuntimeLibrary;
using System.Windows.Forms;
using Ionic.Zip;

namespace ProgramInstaller
{
    static class Program
    {
        public static Form1 form;
        public static string ProgramName = "ShimadzuSecurityConfiguration";

        // The main entry point for the application.
        [STAThread]
        static void Main()
        {
            if (Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName).Length > 1)
            {
                MessageBox.Show("Only one instance of this application is allowed.");
                Environment.Exit(-1);
            }

            // ELoad Embedded Assemblies
            string resource1 = "ProgramInstaller.Ionic.Zip.dll";
            EmbeddedAssembly.Load(resource1, "Ionic.Zip.dll");
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);

            string resource2 = "ProgramInstaller.Interop.IWshRuntimeLibrary.dll";
            EmbeddedAssembly.Load(resource2, "Interop.IWshRuntimeLibrary.dll");
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            form = new Form1();

            Thread t = new Thread(thread);
            t.SetApartmentState(ApartmentState.STA);
            t.Start();

            Application.Run(form);
        }

        private static void thread()
        {
            // Make sure the form has loaded
            while (!form.Created) Thread.Sleep(500);

            // String declarations
            string active_SCU_Directory = @"C:\Shimadzu\Apps\" + ProgramName + @"\";
            string backup_SCU_Directory = @"C:\Shimadzu\Apps\" + ProgramName + @" (Backup)\";

            // Create Directories
            if (!Directory.Exists(active_SCU_Directory)) Directory.CreateDirectory(active_SCU_Directory);
            if (!Directory.Exists(backup_SCU_Directory)) Directory.CreateDirectory(backup_SCU_Directory);

            // Make sure the SCU is not still running
            foreach (string file in Directory.GetFiles(active_SCU_Directory))
            {
                FileInfo fi = new FileInfo(file);
                foreach (var process in Process.GetProcessesByName(fi.Name.Remove(fi.Name.Length-fi.Extension.Length)))
                {
                    process.Kill();
                }
            }

            try
            {
                #region "Commented - Exit if SCU App is running"
                /*bool fileInUse = false;
                foreach (string file in Directory.GetFiles(active_SCU_Directory))
                {
                    try
                    {
                        FileStream fs =
                            System.IO.File.Open(file, FileMode.OpenOrCreate,
                            FileAccess.ReadWrite, FileShare.None);
                        fs.Close();
                    }
                    catch (IOException ex)
                    {
                        fileInUse = true;
                    }
                }
                if (fileInUse == true)
                {
                    MessageBox.Show("Cannot install the Security Configuration Utility while the Security Configuration Tool is running. Please close the SCU Tool before running this installer.");
                    Environment.Exit(-1);
                }*/
                #endregion

                // Delete the backup directory
                if (Directory.Exists(backup_SCU_Directory)) DeleteEntireDirectory(backup_SCU_Directory);

                // Copy Active -> Backup directory
                CopyEntireDirectory(active_SCU_Directory, backup_SCU_Directory);

                // Delete the backed-up active directory
                if (Directory.Exists(active_SCU_Directory)) DeleteEntireDirectory(active_SCU_Directory);

                // Create a new empty active directory
                Directory.CreateDirectory(active_SCU_Directory);
                
                // Extract embedded resource (Program.zip) from this program
                if (System.IO.File.Exists(Path.GetTempPath() + "Program.zip"))
                {
                    System.IO.File.Delete(Path.GetTempPath() + "Program.zip");
                }
                List<string> files = new List<string>();
                files.Add("Program.zip");
                ExtractEmbeddedResource(Path.GetTempPath(), "ProgramInstaller", files);

                // Extract the zip file to the active_SCU_Directory
                using (ZipFile zip = ZipFile.Read(Path.GetTempPath() + "Program.zip"))
                {
                    zip.ExtractAll(active_SCU_Directory);
                }

                // Add shortcut to the Desktop (if it's not already there)
                var startupFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                var shell = new WshShell();
                var shortCutLinkFilePath = startupFolderPath + @"\"+ProgramName+".lnk";
                var windowsApplicationShortcut = (IWshShortcut)shell.CreateShortcut(shortCutLinkFilePath);
                windowsApplicationShortcut.Description = "How to create shortcut for application example";
                windowsApplicationShortcut.WorkingDirectory = active_SCU_Directory;
                windowsApplicationShortcut.TargetPath = active_SCU_Directory + @"ShimadzuSecurityConfiguration.exe";
                windowsApplicationShortcut.Save();

                // Run the new Program
                execute_program_hide_cmd("cmd.exe", "/c \""+ startupFolderPath + "\\"+ProgramName+".lnk\"");

                // Exit the Installer
                Environment.Exit(0);

            }
            catch (Exception ex)
            {
                MessageBox.Show("An error has occurred: " + ex.Message);
                Environment.Exit(-1);
            }
        }

        #region "CMD"
        private static void execute_program_hide_cmd(string filename, string arguments)
        {
            // i.e. execute_program_hide_cmd("C:\\Windows\\Sysnative\\manage-bde.exe", "-resume C:");
            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = filename;
            startInfo.Arguments = arguments;
            process.StartInfo = startInfo;
            process.StartInfo.Verb = "runas";
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.UseShellExecute = false;
            process.Start();
        }
        private static void execute_program_show_cmd(string filename, string arguments)
        {
            // i.e. execute_program_hide_cmd("C:\\Windows\\Sysnative\\manage-bde.exe", "-resume C:");
            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = filename;
            startInfo.Arguments = arguments;
            process.StartInfo = startInfo;
            process.StartInfo.Verb = "runas";
            process.StartInfo.CreateNoWindow = false;
            process.StartInfo.UseShellExecute = true;
            process.Start();
        }
        #endregion
        #region "Copy Entire Directory Methods"
        // Copies a directory and all of it's contents, since there isn't a function in System.IO.Directory to do so.
        public static void CopyEntireDirectory(string sourceDirectory, string targetDirectory)
        {
            DirectoryInfo diSource = new DirectoryInfo(sourceDirectory);
            DirectoryInfo diTarget = new DirectoryInfo(targetDirectory);

            CopyAll(diSource, diTarget);
        }

        public static void CopyAll(DirectoryInfo source, DirectoryInfo target)
        {
            Directory.CreateDirectory(target.FullName);

            // Copy each file into the new directory.
            foreach (FileInfo fi in source.GetFiles())
            {
                //new LogWriter(@"Copying " + target.FullName + "\\" + fi.Name + ": LogPoint#10", "P2_Exceptions.txt");
                fi.CopyTo(Path.Combine(target.FullName, fi.Name), true);
            }

            // Copy each subdirectory using recursion.
            foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
            {
                DirectoryInfo nextTargetSubDir = target.CreateSubdirectory(diSourceSubDir.Name);
                CopyAll(diSourceSubDir, nextTargetSubDir);
            }
        }
        // I made this method because Directory.Delete() doesn't work for directories in the Program Files folder (requires priviledges)
        public static void DeleteEntireDirectory(string directory)
        {
            DeleteAll(new DirectoryInfo(directory));
        }
        public static void DeleteAll(DirectoryInfo source)
        {
            // Delete each file from this directory.
            foreach (FileInfo fi in source.GetFiles())
            {
                System.IO.File.Delete(fi.FullName);
            }

            // Delete each subdirectory using recursion.
            foreach (DirectoryInfo SubDir in source.GetDirectories())
            {
                DeleteAll(SubDir);
            }

            // Delete this Directory Folder
            // Directory.Delete(source.FullName, true);
        }
        #endregion
        #region "Extract Method"
        private static void ExtractEmbeddedResource(string outputDir, string resourceLocation, List<string> files)
        {
            foreach (string file in files)
            {
                using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceLocation + @"." + file))
                {
                    using (FileStream fileStream = new FileStream(System.IO.Path.Combine(outputDir, file), FileMode.Create))
                    {
                        for (int i = 0; i < stream.Length; i++)
                        {
                            fileStream.WriteByte((byte)stream.ReadByte());
                        }
                        fileStream.Close();
                    }
                }
            }
        }
        #endregion
        #region "Assembly Resolver"
        static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            return EmbeddedAssembly.Get(args.Name);
        }
        #endregion
    }
}
