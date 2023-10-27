using Serilog;
using Serilog.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace vs_scanwatch
{

    /// <summary>
    /// vs_scanwatch
    /// The function of this service is to scan the contents of a specified folder and attempt to move it's contents to a subfolder it creates inside this folder
    /// If the subfolder does not exist, the service will create it.
    /// This function is to be executed regularly after a specified amount of time.
    /// 
    /// App.config keys explained
    /// 
    /// "overrideFolderPath" - this key controls the root folder path    
    /// The default value is C:\Scan\.
    /// If a valid path to an existing folder is provided the key will override the default value. 
    /// If a valid path is provided but the folder does not exist, the default root folder path will be used. 
    /// In the case that the default root folder path is used and the C:\Scan\ folder does not exist,
    /// the service will create a C:\Scan\ folder to ensure operation.
    /// Values of "false", "False" or "FALSE" will instruct the app to use the default value.
    ///  
    /// "overrideDate" - this key controls the output folder behaviour
    /// The default behaviour is to set the folder to the current date in Belgrade in yyyyMMdd format
    /// If a valid date in the format yyyyMMdd is provided, the key date will be used instead
    /// 
    /// "scanInterval" - this key controls the time interval that passes before the service attempts to execute it's main programming again
    /// The default value is 30000. The value represents miliseconds, so 30000 equals to a wait time of 30 seconds.
    /// The value entered needs to be a valid number or else the default value will be used.
    /// The minimal wait time is 10 seconds(10000), if a lower value is entered, the default value will be used.
    /// Values of "false", "False" or "FALSE" will instruct the app to use the default value.
    /// 
    /// "debug" - this key controls wether debug logging is enabled or disabled
    /// Values of "true", "True" or "TRUE" will enable debug logging inside a separate debug log file instead of the regular file.
    /// </summary>
    public class CopyQueue //This class exists just to ensure that the copy operations get performed sequentially and thus not cause a strain on the system's IO resources
    {        
        public int error;
        public FileStream[] fileStreams;    //TODO dispose of array resources on each run
                                            //fileStreams need to be done through an instance of a separate class to ensure that they get properly disposed once a successful copy has been confirmed
    }
    static class FolderPath //This class exists just to pass the App.config settings from the OnStart to the OnElapsedTime method
    {
        public static string rootDirectoryPath; //this needs to be a valid path to an existing folder, if path is not valid/folder does not exist default value is used
                                                //default value is "C:\Scan\" and if that folder doesn't exist "C:\Scan\" will be created
        public static string outputFolderPath;  //this needs to be formated as yyyyMMdd and needs to be a valid date to be accepted
                                                //default value is current date in Belgrade formated as yyyyMMdd
    }
    public partial class Service1 : ServiceBase
    {
        Timer timer = new Timer();
        public Service1()
        {
            InitializeComponent();
        }
        
        protected override void OnStart(string[] args)
        {            
            string overrideFolderPath = ConfigurationManager.AppSettings["overrideFolderPath"];    //This can be loaded, but has a default value
            if ("false" == overrideFolderPath || "FALSE" == overrideFolderPath || "False" == overrideFolderPath)
            {
                overrideFolderPath = default(string);
            }

            string overrideDate = ConfigurationManager.AppSettings["overrideDate"];
            FolderPath.rootDirectoryPath = SetRootDirectory(overrideFolderPath);

            string debug = ConfigurationManager.AppSettings["debug"];

            string logName = "\\log\\LogFile_Debug_" + TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, TimeZoneInfo.FindSystemTimeZoneById("Central Europe Standard Time")).ToString("yyyy.MM.dd") + ".log";

            if ("true" != debug && "TRUE" != debug && "True" != debug)
            {
                logName = "\\log\\LogFile_" + TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, TimeZoneInfo.FindSystemTimeZoneById("Central Europe Standard Time")).ToString("yyyy.MM.dd") + ".log";
            }
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.File(FolderPath.rootDirectoryPath + "\\log\\LogFile_Error.txt", restrictedToMinimumLevel: Serilog.Events.LogEventLevel.Warning)
                .WriteTo.File(FolderPath.rootDirectoryPath + logName, restrictedToMinimumLevel: Serilog.Events.LogEventLevel.Debug)
                .CreateLogger();

            if ("true" != debug && "TRUE" != debug && "True" != debug)
            {
                Log.Logger = new LoggerConfiguration()
                    .WriteTo.Sink((ILogEventSink)Log.Logger)
                    .WriteTo.File(FolderPath.rootDirectoryPath + logName, restrictedToMinimumLevel: Serilog.Events.LogEventLevel.Information)
                    .CreateLogger();
            }

            Log.Information($"Service started at system time {DateTime.Now}.");

            Log.Debug("Debug logging enabled.");    //even though this gets executed either way, this only shows up in logs if debug logging is enabled

            string folderNameDate = GetFolderName(overrideDate);
            FolderPath.outputFolderPath = SetOutputDirectory(FolderPath.rootDirectoryPath, folderNameDate);


            int timerInterval = 30000;   //This can be loaded, but has a default value of 30 seconds, value is in miliseconds
            string timerConfig = ConfigurationManager.AppSettings["scanInterval"];
            if ("false" != timerConfig && "FALSE" != timerConfig && "False" != timerConfig)
            {
                Int32.TryParse(timerConfig, out timerInterval);

                if (timerInterval < 10000)  //Minimal interval time is 10 seconds, if we have an invalid value or a value of less than 10 seconds, the default value of 30 seconds will be selected
                {
                    timerInterval = 30000;
                }
            }

            Log.Information($"Service interval set at {timerInterval} miliseconds.");
            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            timer.Interval = timerInterval;
            timer.Enabled = true;
        }

        protected override void OnStop()
        {
            Log.Information($"Service stopped at system time {DateTime.Now}.");
        }        

        //This method computes the MD5 hash hash of the provided file, attempts to copy it to the provided destination folder with a unique new name, compares the hashes of the new and old file and if matching deletes the old file, if not deletes the new file and unlocks the old file 
        public static void CopyCompareAndDelete(string inputFileName, string destination, long index, CopyQueue cq, long time)
        {
            string outputFileName = destination + (time + index).ToString() + Path.GetExtension(cq.fileStreams[index].Name);  //Constructing the new/output file path, along with name and file extension
            try
            {                
                string originalHash = ComputeFileHash(cq.fileStreams[index]); //Computing the original file hash, since the first thing we do is lock the file by opening a filestream, the hash needs to be computed from the filestream            
                Log.Debug($"Attempting copy operation: \"{inputFileName}\" -> \"{outputFileName}\".");

                string copyHash;

                if (File.Exists(outputFileName))    //If a file with the same filename already exists it will be skipped and unlocked
                {
                    throw new Exception("Filename already exists.");
                }

                //This code block copies the input file to the specified destination
                using (FileStream fs2 = File.Create(outputFileName))
                {
                    cq.fileStreams[index].Position = 0;   //IMPORTANT, will not work without this, this ensures that we're reading the filestream of the original file from the begining, the position was at the end because we computed the file hash earlier
                    cq.fileStreams[index].CopyTo(fs2);
                    fs2.Flush();    //IMPORTANT, will not work without this, this ensures that we're writing all the bytes to the new file before closing the filestream
                }

                using (FileStream fs2 = new FileStream(outputFileName, FileMode.Open, FileAccess.Read, FileShare.None)) //since the flush() method closes the filestream, we need to open a new one to compute the hash
                {
                    copyHash = ComputeFileHash(fs2);
                }


                //This block compares the hashes of the input and output file and throws an exception if it detects a mismatch
                if (originalHash != copyHash)
                {
                    File.Delete(outputFileName);    //Deletes the output file created because it is not a match, this needs to be here instead of in the catch block because if we run into an existing file with the same name it also throws an exception, so if this line were in the catch block it would delete the file in that case
                    throw new Exception("File Hash of copy does not match original");
                }

                Log.Information($"Copy operation success: \"{inputFileName}\" -> \"{outputFileName}\".");

                //Disposing of input file stream and deleting the input file
                cq.fileStreams[index].Dispose();
                File.Delete(inputFileName);
                Log.Debug($"Deleted file: \"{inputFileName}\".");
            }
            catch (Exception e)
            {
                cq.fileStreams[index].Dispose();
                Log.Error(e, $"Copy operation failed: \"{inputFileName}\" -> \"{outputFileName}\". ");
                cq.error++;    //Error counter increases, used just for a debug message
                Log.Debug($"Error count set  to {cq.error}.");
            }
            return;
        }

        //This method computes the MD5 hash of a file
        public static string ComputeFileHash(FileStream stream)
        {
            stream.Position = 0;
            MD5 md5Service = MD5.Create();
            byte[] hash = md5Service.ComputeHash(stream);
            md5Service.Dispose();

            string hashString = BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
            Log.Debug($"MD5 Hash of {stream.Name} calculated: \"{hashString}\".");

            return hashString;
        }

        //This method checks to see if the provided folder exists on the drive before selecting it and if not, it selects the default folder
        public static string SetRootDirectory(string rootFolder)
        {
            if (Directory.Exists(rootFolder))
            {
                //This does nothing because it gets called before the logger gets initialized, so it's commented out and left for reference
                //Log.Information($"Override folder found on drive. Root folder set to \"{rootFolder}\".");
                return rootFolder;
            }

            string defaultFolder = @"C:\Scan\";

            if (!Directory.Exists(defaultFolder))
            {
                System.IO.Directory.CreateDirectory(defaultFolder);
            }

            //This does nothing because it gets called before the logger gets initialized, so it's commented out and left for reference
            //Log.Information($"Override folder not found on drive. Root folder set to default value \"{defaultFolder}\".");
            return defaultFolder;
        }

        //This method checks to see if a subfolder with the selected date name already exists inside the specified root folder and if not, it creates that folder
        public static string SetOutputDirectory(string rootFolder, string dateFolder)
        {
            string outputFolder = Path.Combine(rootFolder, dateFolder)+"\\";
            System.IO.Directory.CreateDirectory(outputFolder);
            return outputFolder;
        }


        //This method checks to see if a valid date override has been entered and returns the name of the output folder
        public static string GetFolderName(string dateTimeOverride)
        {
            Log.Information($"Root folder set to \"{FolderPath.rootDirectoryPath}\"."); //This gets logged here because the root folder gets selected before the logger gets initialized
            DateTime validDate;
            //Checks if the provided override is a valid date and if not returns current date in Belgrade
            if (false == DateTime.TryParseExact(dateTimeOverride, "yyyyMMdd", new CultureInfo("sr-Latn-RS"), DateTimeStyles.AdjustToUniversal, out validDate))
            {
                return SetFolderNameToBelgradeTimeDate();
            }

            string output = validDate.ToString("yyyyMMdd");

            Log.Information($"Valid override date detected. Destination folder set to \"{output}\".");
            return output;
        }

        //This method gets called if no valid date override has been provided and generates the folder name according to the current date in Belgrade
        public static string SetFolderNameToBelgradeTimeDate()
        {
            string output = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, TimeZoneInfo.FindSystemTimeZoneById("Central Europe Standard Time")).ToString("yyyyMMdd");
            Log.Information($"No valid override date detected. Destination folder set to \"{output}\".");
            return output;
        }

        //This is the main method that the service runs every time the timer passes the provided time interval(or if the provided interval isn't a valid number higher than 10 seconds, it runs every 30 seconds by default)
        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            string[] files = Directory.GetFiles(FolderPath.rootDirectoryPath);  //Gets a list of all files in the selected directory
            bool[] failedStreams = new bool[files.Length];
            Log.Debug($"Found {files.Length} files.");
            if (0 != files.Length)
            {                
                long milliseconds = DateTimeOffset.Now.ToUnixTimeMilliseconds();    //This is needed to ensure all output files have unique names            
                Log.Debug($"Miliseconds is {milliseconds}.");

                CopyQueue copyQueue = new CopyQueue();
                copyQueue.fileStreams = new FileStream[files.Length];
                copyQueue.error = 0;    //Error counter for log
                Log.Debug($"Error count set to {copyQueue.error}.");

                for (int i = 0; i < files.Length; i++)
                {
                    failedStreams[i] = false;
                }

                //TODO - isolate Parallel.ForEach into a separate method called LockAndQueueFiles
                //We're using Parallel.ForEach to lock all the files at the same time and prevent them from being accessed by another instance of this program by accident
                Parallel.ForEach(files, (file, pls, fileIndex) =>
                {
                    //We provide each thread the path to the file it needs to attempt to copy, the output file path, the index of the file in the original file list, an instance of the CopyQueue class that is needed to ensure that files are copied sequentially, and the timestamp for unique naming of output files
                    //CopyCompareAndDelete(file, FolderPath.outputFolderPath, fileIndex, copyQueue, milliseconds);  //Doesn't work with larger files(missed this by only testing it with 10kb doc files)
                    
                    try
                    {
                        copyQueue.fileStreams[fileIndex] = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.None); //The first thing we do is locking the file to ensure it doesn't get accessed by for example another instance of the service during the copy attempt
                        Log.Debug($"Opened fileStream[{fileIndex}] for {file} successfully.");
                    } catch
                    {
                        failedStreams[fileIndex] = true;
                        copyQueue.fileStreams[fileIndex].Dispose();
                        Log.Error($"Failed to create FileStream for {file}.");
                        copyQueue.error++;  //Error counter increases, used just for a debug message
                        Log.Debug($"Error count set to {copyQueue.error}.");
                    }                    
                });

                for(int i=0; i < files.Length; i++)
                {
                    if (false == failedStreams[i])
                    {
                        Log.Debug($"failedStreams[{i}] is false, Calling CopyCompareAndDelete.");
                        CopyCompareAndDelete(files[i], FolderPath.outputFolderPath, i, copyQueue, milliseconds);
                    }
                }

                if (0 == copyQueue.error)
                {
                    Log.Debug($"All files found in \"{FolderPath.rootDirectoryPath}\" successfully moved to \"{FolderPath.outputFolderPath}\".");
                }
                else
                {
                    Log.Debug($"Not all files found in \"{FolderPath.rootDirectoryPath}\" successfully moved to \"{FolderPath.outputFolderPath}\". There were {copyQueue.error} errors");
                }
            }
            else
            {
                Log.Debug($"No files found in: \"{FolderPath.rootDirectoryPath}\".");
            }
        }
    }
}
