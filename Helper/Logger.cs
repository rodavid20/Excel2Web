using System;
using System.Diagnostics;
using System.IO;
using System.Configuration;

namespace Excel2Web.Helper
{
    public static class Logger
    {
        private static string fileLocation;

        static Logger()
        {
            //Change this value in the web.config to correct log folder on the server 
            //and give write permissions
            fileLocation = ConfigurationManager.AppSettings["LogLocation"];
        }

        
        public static void WriteToLog(string logInformation, EventLogEntryType eType)
        {
            if (string.IsNullOrEmpty(logInformation)) return;

            // Create the subfolder
            Directory.CreateDirectory(fileLocation);
            //generate random number for unique log file name
            Random r = new Random();
            
            string fileName = Path.Combine(fileLocation, String.Format("log{0}{1}.txt", 
                    DateTime.Today.ToString("ddMMyyyy"), r.Next(1000)));

            // Create a writer and open the file:
            StreamWriter log = new StreamWriter(fileName);

            // Write to the file:
            log.WriteLine(DateTime.Now + " " + eType);
            log.WriteLine(logInformation);
            // Close the stream:
            log.Close();

        }
    }
}