using CGCDReportParser.ViewModels;
using CGCDReportParser.Views;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace CGCDReportParser.Models
{
    internal class deps
    {

        public static bool IsLibreOfficeWriterAvailable()
        {
            try
            {
                // Create a new process for dpkg.
                Process dpkgProcess = new Process();

                // Assign dpkg as the process name.
                dpkgProcess.StartInfo.FileName = "dpkg";

                // Provide arguments to dpkg to check for libreoffice-writer
                dpkgProcess.StartInfo.Arguments = "--status libreoffice-writer";

                // Make the process window hidden.
                dpkgProcess.StartInfo.CreateNoWindow = true;
                dpkgProcess.StartInfo.UseShellExecute = false;

                // Capture the output
                dpkgProcess.StartInfo.RedirectStandardOutput = true;

                // Start the process.
                dpkgProcess.Start();

                // Read the output to a string
                string output = dpkgProcess.StandardOutput.ReadToEnd();

                // If the process started successfully, kill it immediately.
                dpkgProcess.Kill();

                // Check the output for the status of libreoffice-writer
                if (output.Contains("install ok installed"))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception)
            {
                // If an exception has been thrown, dpkg couldn't be started or libreoffice-writer is not installed.
                return false;
            }
        }
        public static bool IsSOfficeAvailable()
        {
            try
            {
                var pathVar = Environment.GetEnvironmentVariable("PATH");

                Debug.WriteLine(pathVar);
                // Create a new process for soffice.
                Process sofficeProcess = new Process();

                // Assign soffice as the process name.
                sofficeProcess.StartInfo.FileName = "soffice";

                // Make the process window hidden.
                sofficeProcess.StartInfo.CreateNoWindow = true;
                sofficeProcess.StartInfo.UseShellExecute = false;

                // Start the process.
                sofficeProcess.Start();

                // If the process started successfully, kill it immediately.
                sofficeProcess.Kill();

                //linux needs to check if writer is also installed alongside da commonz
                if (Environment.OSVersion.Platform == PlatformID.Unix)
                {
                    if (IsLibreOfficeWriterAvailable())
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                return true;
            }
            catch (Exception)
            {
                // If an exception has been thrown, soffice couldn't be started.
                return false;
            }
        }

        static void AddToUserAndProcessPath()
        {
            string libreOfficePath = @"C:\Program Files\LibreOffice\program\";

            // Add to the User PATH
            string userPathVar = Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.User);
            if (userPathVar != null && !userPathVar.Contains(libreOfficePath))
            {
                string newUserPathValue = $"{userPathVar};{libreOfficePath}";
                Environment.SetEnvironmentVariable("PATH", newUserPathValue, EnvironmentVariableTarget.User);
            }


            // Add to the Process PATH
            string processPathVar = Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.Process);
            if (processPathVar != null && !processPathVar.Contains(libreOfficePath))
            {
                string newProcessPathValue = $"{processPathVar};{libreOfficePath}";
                Environment.SetEnvironmentVariable("PATH", newProcessPathValue, EnvironmentVariableTarget.Process);
            }
        }
        public static void InstallDependencies()
        {

            if (Environment.OSVersion.Platform == PlatformID.Unix)
            {

                ProcessStartInfo checkStartInfo = new ProcessStartInfo()
                {
                    FileName = "dpkg",
                    Arguments = "-s libreoffice-common libreoffice-writer",
                    RedirectStandardError = false, // dpkg sends its output to stderr
                    UseShellExecute = false,
                    CreateNoWindow = true,
                };

                Process checkProcess = new Process() { StartInfo = checkStartInfo };
                checkProcess.Start();

                string output = checkProcess.StandardError.ReadToEnd();
                checkProcess.WaitForExit();

                if (output.Contains("is not installed"))
                {

                    ProcessStartInfo installStartInfo = new ProcessStartInfo()
                    {
                        FileName = "sudo",
                        Arguments = "apt-get install -y libreoffice-common libreoffice-writer",
                        UseShellExecute = false,
                        CreateNoWindow = true,
                    };

                    Process installProcess = new Process() { StartInfo = installStartInfo };
                    installProcess.Start();
                    installProcess.WaitForExit();
                }
            }
            else if (Environment.OSVersion.Platform == PlatformID.Win32NT)
            {
                string downloadUrl = "https://download.documentfoundation.org/libreoffice/stable/7.5.5/win/x86_64/LibreOffice_7.5.5_Win_x86-64.msi";
                string installerPath = Path.Combine(Path.GetTempPath(), "LibreOffice_7.5.5_Win_x86-64.msi");

                // Download the installer
                WebClient webClient = new WebClient();
                webClient.DownloadFile(downloadUrl, installerPath);

                // Run the installer
                Process process = new Process();
                process.StartInfo.FileName = "msiexec.exe"; // Change this to msiexec.exe
                process.StartInfo.Arguments = $"/i \"{installerPath}\" /passive"; // Update the arguments to include the installer path
                process.Start();
                process.WaitForExit();

                AddToUserAndProcessPath();
            }
        }
        /*
        public static void InstallDocx2Pdf()
        {
            ProcessStartInfo start = new ProcessStartInfo();
            start.FileName = "python";
            start.Arguments = "-m pip install docx2pdf";
            start.UseShellExecute = false; //no system shell.
            start.RedirectStandardOutput = true; // Any output
            start.RedirectStandardError = true; // Any error 
            start.CreateNoWindow = true; // Don't create a new window

            using (Process process = Process.Start(start))
            {
                string stderr = process.StandardError.ReadToEnd();
                process.WaitForExit();
            }
        }
        
        public static bool IsDocx2PdfInstalled()
        {
            ProcessStartInfo start = new ProcessStartInfo();
            start.FileName = "python";
            start.Arguments = "-m pip show docx2pdf";
            start.UseShellExecute = false; // Do not use OS shell
            start.RedirectStandardOutput = true; // Any output, true  
            start.RedirectStandardError = true; // Any error, true 
            start.CreateNoWindow = true; // Don't create a new window

            using (Process process = Process.Start(start))
            {
                using (StreamReader reader = process.StandardOutput)
                {
                    string result = reader.ReadToEnd(); // Read the output: if it is empty, the package is not installed
                    process.WaitForExit();
                    return !string.IsNullOrEmpty(result);
                }
            }
        }
        */


    }
}
