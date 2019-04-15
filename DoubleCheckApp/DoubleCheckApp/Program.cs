using System;
using System.Diagnostics;
using System.IO;
using CoreySutton.Xrm.Utilities;
using Microsoft.Xrm.Sdk;

namespace CoreySutton.Xrm.DoubleCheckApp
{
    class Program
    {
        private const string _break = "#########################################################";

        static void Main(string[] args)
        {
            try
            {
                IOrganizationService organizationService = GetOrganizationService(args[0]);
                LaunchPackageDeploymentExe();
            }
            catch (Exception ex)
            {
                WriteException(ex);
            }

            Console.WriteLine();
            Console.WriteLine("Execution finished. Press <enter> to close...");
            Console.ReadLine();
        }

        static IOrganizationService GetOrganizationService(string connectionString = null)
        {
            IAppSettingsManager appSettingsManager = new AppSettingsManager(Properties.Settings.Default);
            ICrmCredentialManager crmCredentialManager = new CrmCredentialManager(appSettingsManager);
            IOrganizationService organizationService = CrmConnectorUtil.Connect(crmCredentialManager);

            return organizationService;
        }

        static void LaunchPackageDeploymentExe()
        {
            string workingDirectory = Environment.CurrentDirectory;
            string solutionPackagerPath = Path.Combine(workingDirectory, @"..\coretools\SolutionPackager.exe");
            string solutionZipPath = Path.Combine(workingDirectory, @"BizDev_1_0_0_0.zip");
            string extractedSolutionPath = Path.Combine(workingDirectory, @"solutions\BizDev_1_0_0_0");
            // Off/Error/Warning/Info/Verbose
            string errorLevel = "Verbose";
            string arguments = $"/action Extract /zipfile {solutionZipPath} /folder {extractedSolutionPath} /errorlevel {errorLevel}";

            Console.WriteLine($"{nameof(workingDirectory)}:\t{workingDirectory}");
            Console.WriteLine($"{nameof(solutionPackagerPath)}:\t{solutionPackagerPath}");
            Console.WriteLine($"{nameof(solutionZipPath)}:\t{solutionZipPath}");
            Console.WriteLine($"{nameof(extractedSolutionPath)}:\t{extractedSolutionPath}");

            Console.WriteLine();
            Console.WriteLine(_break);
            Console.WriteLine("SolutionPackager.exe");
            Console.WriteLine(_break);
            Console.WriteLine();

            Console.WriteLine($"Starting with arguments: {arguments}");
            Console.WriteLine();

            // Use ProcessStartInfo class
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                CreateNoWindow = false,
                UseShellExecute = false,
                FileName = solutionPackagerPath,
                WindowStyle = ProcessWindowStyle.Hidden,
                Arguments = arguments,
                WorkingDirectory = workingDirectory,
            };

            try
            {
                // Start the process with the info we specified.
                // Call WaitForExit and then the using statement will close.
                using (Process exeProcess = Process.Start(startInfo))
                {
                    exeProcess.WaitForExit();
                }
            }
            catch (Exception ex)
            {
                WriteException(ex);
            }
        }

        static void WriteException(Exception ex)
        {
            Console.WriteLine();
            Console.WriteLine(_break);
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write("An error occured: ");
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.Write(ex.Message);
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine(ex.StackTrace);
            Console.ForegroundColor = ConsoleColor.Gray;

            int count = 1;
            Exception innerException = ex.InnerException;
            while (innerException != null)
            {
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Inner Exception ({count}):");
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine(ex.InnerException.Message);
                Console.ForegroundColor = ConsoleColor.DarkGray;
                Console.WriteLine(ex.InnerException.StackTrace);
                Console.ForegroundColor = ConsoleColor.Gray;

                innerException = innerException.InnerException;
                count++;
            }

            Console.WriteLine(_break);
        }
    }
}
