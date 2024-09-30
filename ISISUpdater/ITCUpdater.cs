using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ITCUpdater
{
    public enum AppName { Unspecified, FrontEnd, SurveyReport };
    class ITCUpdater
    {

        string appUpdateFolder;// = "\\\\psychfile\\psych$\\psych-lab-gfong\\SMG\\Access\\Application Updates";
        string appInstallFolder; //= "D:\\users\\ITC ISIS";
        

        public ITCUpdater(string updateFolder, string installFolder)
        {
            appUpdateFolder = updateFolder;
            appInstallFolder = installFolder;
        }

        public int Install()
        {
            try
            {
                string fileName;
               
                // install FE
                fileName = "ISIS FrontEnd Ver." + LatestFE() + ".accdb";
                File.Copy(appUpdateFolder + "\\" + fileName, appInstallFolder + "\\" + fileName);

                // install SR
                fileName = "ISIS ISR Ver." + LatestSR() + ".accdb";
                File.Copy(appUpdateFolder + "\\" + fileName, appInstallFolder + "\\" + fileName);
            }
            catch
            {
                return 1;
            }
            return 0;
        }

        public int Install(AppName app)
        {
            string fileName;

            try
            {
                switch (app)
                {
                    case AppName.FrontEnd:
                        // install FE
                        fileName = "SDI FrontEnd Ver." + LatestFE() + ".accdb";
                        File.Copy(appUpdateFolder + "\\" + fileName, appInstallFolder + "\\" + fileName);
                        break;

                    case AppName.SurveyReport:
                        // install SR
                        fileName = "SDI ISR Ver." + LatestSR() + ".accdb";
                        File.Copy(appUpdateFolder + "\\" + fileName, appInstallFolder + "\\" + fileName);
                        break;
                    case AppName.Unspecified:
                        Install();
                        break;
                }
            }
            catch
            {
                return 1;
            }
            return 0;
        }

        public bool HaveLatestVersionFE()
        {
            Version latestFE = LatestVersionInFolder(appUpdateFolder, AppName.FrontEnd);

            Version currentFE = LatestVersionInFolder(appInstallFolder, AppName.FrontEnd);

            return currentFE == latestFE;
        }

        public bool HaveLatestVersionSR()
        {
            

            Version latestSR = LatestVersionInFolder(appUpdateFolder, AppName.SurveyReport);

            Version currentSR = LatestVersionInFolder(appInstallFolder, AppName.SurveyReport);

            return currentSR == latestSR;
        }

        public Version LatestVersionInFolder(string folder, AppName app)
        {

            DirectoryInfo d = new DirectoryInfo(folder);
            FileInfo[] files = d.GetFiles();
            string ver = "";
      
            Version currentVersion = new Version(0,0,0);
            int verLocation; // location of Ver. in the filename
            int extLocation; // location of accdb in the filename
            foreach (FileInfo f in files)
            {

                switch (app)
                {
                    case AppName.FrontEnd:
                        if (!f.Name.Contains("SDI FrontEnd") || f.Name.Contains("beta"))
                            continue;
                        break;
                    case AppName.SurveyReport:
                        if (!f.Name.Contains("SDI ISR") || (f.Name.Contains("Lite")))
                            continue;
                        break;
                }


                verLocation = f.Name.IndexOf("Ver.");
                extLocation = f.Name.IndexOf(".accdb");

                if (verLocation == -1 || extLocation == -1)
                    continue;

                ver = f.Name.Substring(verLocation + "Ver.".Length, extLocation - verLocation - "Ver.".Length);

                //if (!Double.TryParse(ver, out double result))
                 //   continue;

                string[] digits = ver.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
                Version v;
                if (digits.Length ==1 )
                    v = new Version(Int32.Parse(digits[0]), 0, 0);
                else if (digits.Length == 2)
                    v = new Version(Int32.Parse(digits[0]), Int32.Parse(digits[1]), 0);
                else
                    v = new Version(Int32.Parse(digits[0]), Int32.Parse(digits[1]), Int32.Parse(digits[2]));

                if (IsNewer(v, currentVersion))
                {
                    currentVersion = v;
                }

            }
            return currentVersion;
        }

        public bool IsNewer(Version v1, Version v2)
        {
            if (v1.Major > v2.Major)
                return true;

            if (v1.Major < v2.Major)
                return false;

            if (v1.Minor > v2.Minor)
                return true;

            if (v1.Minor < v2.Minor)
                return false;

            if (v1.Build > v2.Build)
                return true;

            if (v1.Build < v2.Build)
                return false;

            return true;
        }

        public string LatestFE()
        {
            Version v = LatestVersionInFolder(appUpdateFolder, AppName.FrontEnd);

            return ConvertVersionNumber(v);

        }

        public string LatestSR()
        {
            Version v = LatestVersionInFolder(appUpdateFolder, AppName.SurveyReport);
            return ConvertVersionNumber(v); 

        }

        public string ConvertVersionNumber(Version v)
        {
            string result;

            if (v.Minor <= 0)
                result = v.ToString(1);
            else if (v.Build <= 0)
                result = v.ToString(2);
            else
                result = v.ToString(3);

            return result;
        }
    }
}
