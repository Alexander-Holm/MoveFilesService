using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace MoveFilesService
{
    public partial class MoveFilesService : ServiceBase
    {
        // https://stackoverflow.com/questions/19151363/windows-service-to-run-a-function-at-specified-time

        Timer timer;
        DateTime scheduledTime;
        string baseFolder;
        string targetFolder;
        string logfile;
        string readme;

        public MoveFilesService()
        {
            InitializeComponent();

            timer = new Timer();
            // Tiden som filerna flyttas
            scheduledTime = DateTime.Today.AddHours(13).AddMinutes(30);
            
            baseFolder = @"C:\AlexanderHolm\MoveFilesService\";
            targetFolder = baseFolder + @"TargetFolder\";
            logfile = baseFolder + "@log.txt";
            readme = baseFolder + "@readme.txt";
        }

        protected override void OnStart(string[] args)
        {
            Setup();
            WriteToLog("Service startad");

            // Tid tills scheduleTime idag, timer.Interval vill ha det i millisekunder.
            double firstInterval = scheduledTime.Subtract(DateTime.Now).TotalMilliseconds;            
            if (firstInterval < 0)
                // Om scheduleTime redan har varit idag sätts timern för imorgon
                firstInterval += TimeSpan.FromDays(1).TotalMilliseconds;

            timer.Interval = firstInterval;
            timer.Elapsed += new ElapsedEventHandler(OnTimer);
            timer.Enabled = true;
        }

        protected override void OnStop()
        {
            Setup();
            WriteToLog("Service stoppad");
        }

        protected void OnTimer(object sender, ElapsedEventArgs e){
            Setup();
            // Lägg till en dag på timern
            timer.Interval = TimeSpan.FromDays(1).TotalMilliseconds;

            string[] excludedFiles = new string[] { readme, logfile };
            string[] filesToMove = Directory.GetFiles(baseFolder).Except(excludedFiles).ToArray();
            if(filesToMove.Length > 0)
            {
                MoveFiles(filesToMove);
                WriteToLog(filesToMove.Length +" filer flyttade");
            }                 
        }

        private void MoveFiles(string[] filesToMove)
        {
            Directory.CreateDirectory(targetFolder);
            foreach (var filepath in filesToMove)
            {
                string newFilepath = Path.Combine(targetFolder, Path.GetFileName(filepath));
                // File.Move har inte overwrite. I nyare versioner av .Net Core finns File.Move med overwrite.
                // Här används Copy med Delete istället för Delete sen Move,
                // för att inte förlora filer om något går fel.
                try
                {
                    // Overwrite == true
                    File.Copy(filepath, newFilepath, true);
                    File.Delete(filepath);
                }
                catch (Exception e)
                {
                    WriteToLog("Error: " + e.Message);
                }
            }
        }

        private void WriteToLog(string logMessage)
        {
            File.AppendAllText(logfile, DateTime.Now + " - " + logMessage + "\n");
        }

        private void Setup()
        {
            // Skapar även baseFolder
            Directory.CreateDirectory(targetFolder);

            if (!File.Exists(readme))
            {
                using (var sw = File.CreateText(readme))
                {
                    sw.WriteLine("Kl." + scheduledTime.TimeOfDay + " flyttas alla filer till mappen TargetFolder (inte @readme.txt och @log.txt)");
                }
            }
            if (!File.Exists(logfile))
            {
                WriteToLog("Logfil skapad");
            }

        }
    }
}
