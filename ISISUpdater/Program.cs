using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Access = Microsoft.Office.Interop.Access;

namespace ITCUpdater
{
    class Program
    {

        static string appUpdateFolder = "\\\\psychfile\\psych$\\psych-lab-gfong\\SMG\\Access\\Application Updates";
        static string appInstallFolder = "D:\\users\\ITC SDI";

        static void Main(string[] args)
        {
            ITCUpdater updater = new ITCUpdater(appUpdateFolder, appInstallFolder);

            // process args to determine which apps to install
            AppName app = AppName.Unspecified;
            string filename = appInstallFolder + "\\SDI FrontEnd Ver.";

            if (args.Length != 0)
            {
                switch (args[0])
                {
                    case "f":
                        app = AppName.FrontEnd;
                        filename = appInstallFolder + "\\SDI FrontEnd Ver.";
                        break;
                    case "r":
                        app = AppName.SurveyReport;
                        filename = appInstallFolder + "\\SDI ISR Ver.";
                        break;
                }
            }


            // if directory doesn't exist, install apps, launch
            if (!Directory.Exists(appInstallFolder))
            {
                int result = 0;
                Directory.CreateDirectory(appInstallFolder);
                               
                result = updater.Install(app);
                
                //// launch apps when install is complete
                //if (result == 0)
                //{
                //    // launch Access, look at args and open FE/ISR depending on args
                //}
            }
            else // if directory does exist, compare versions, update, launch
            {
               
                int result=0;
                if (args.Length == 0)
                {
                    if (!updater.HaveLatestVersionFE())
                    {
                        result = updater.Install(AppName.FrontEnd);
                    }

                    if (result == 1)
                    {
                        Console.WriteLine("Error installing FrontEnd.");
                    }
                    if (!updater.HaveLatestVersionSR())
                    {
                        result = updater.Install(AppName.SurveyReport);
                    }

                    if (result == 1)
                    {
                        Console.WriteLine("Error installing Survey Report.");
                    }
                }
                else
                {
                    switch (args[0])
                    {
                        case "f":
                            if (!updater.HaveLatestVersionFE())
                            {
                                result = updater.Install(AppName.FrontEnd);
                            }

                            if (result == 1)
                            {
                                Console.WriteLine("Error installing FrontEnd.");
                            }
                            break;
                        case "r":
                            if (!updater.HaveLatestVersionSR())
                            {
                                result = updater.Install(AppName.SurveyReport);
                            }

                            if (result == 1)
                            {
                                Console.WriteLine("Error installing Survey Report.");
                            }
                            break;
                    }
                }


                

                
            }

            
            Version latest = updater.LatestVersionInFolder(appInstallFolder, app);

            filename += updater.ConvertVersionNumber(latest) + ".accdb";
            
            Access.Application oAccess = null;
            // Start a new instance of Access for Automation:
            oAccess = new Access.Application();

            // Open a database in exclusive mode, with low security
            oAccess.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityLow;
            oAccess.OpenCurrentDatabase(filename, true);
            oAccess.DoCmd.Maximize();
            
            oAccess.Visible = true;
        }
        
    }
}
