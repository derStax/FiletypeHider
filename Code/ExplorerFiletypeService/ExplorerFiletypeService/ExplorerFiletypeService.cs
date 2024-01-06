using System;
using System.IO;
using SHDocVw;
using System.ServiceProcess;
using System.Collections.Generic;
using System.Timers;

namespace ExplorerFiletypeService {
    public partial class ExplorerFiletypeService : ServiceBase {

        private readonly Timer _timer = new Timer();
        public ExplorerFiletypeService() {
            InitializeComponent();

            ServiceName = "Explorer Filetype Extention Manager";
            this.CanHandlePowerEvent = true;
            this.CanHandleSessionChangeEvent = true;
            this.CanPauseAndContinue = true;
            this.CanShutdown = true;
            this.CanStop = true;

            _timer.Elapsed += new ElapsedEventHandler(HideFiles);
            _timer.Interval = 5000; //time in milliseconds to check for given filetypes
            _timer.Enabled = true;
        }

        protected override void OnStart(string[] args) {
            /* Create important Files if they dont exist yet. */
            CreateFile(HiddenFilesPath);
            CreateFile(ExtentionToHidePath);
            CreateFile(LogPath);

            _extentionsToHide = File.ReadAllLines(ExtentionToHidePath); // read the hiddenFiletypes list

            ShowFiles(); // show all files which got hidden "last time"
        }

        // paths, look like:
        // C:\Users\Public\Documents\FourteenDynamics\ExplorerFiletypeHandler\[FILE]
        private readonly string HiddenFilesPath = 
            Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) + "\\FourteenDynamics\\ExplorerFiletypeHandler\\HiddenFiles.txt";
        private readonly string ExtentionToHidePath =
            Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) + "\\FourteenDynamics\\ExplorerFiletypeHandler\\ExtentionsToHide.txt";
        private readonly string LogPath =
            Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) + "\\FourteenDynamics\\ExplorerFiletypeHandler\\Logs.txt";


        /// <summary>
        /// Saves the extentions to hide (example: *.meta) - only reads when starting the service!
        /// </summary>
        private string[] _extentionsToHide = new string[] { };

        /// <summary>
        /// Gets called from timer, gets all files in currently opened explorers and hides the " <see cref="_extentionsToHide"/> " inside. 
        /// </summary>
        private void HideFiles(object source, ElapsedEventArgs e) {
            foreach(FileInfo f in GetFiles()) {
                if(f.Attributes != FileAttributes.Hidden) {
                    f.Attributes = FileAttributes.Hidden;
                    if(!File.ReadAllText(HiddenFilesPath).Contains(f.FullName)) {
                        if(!File.Exists(HiddenFilesPath)) {
                            // Create a file to write to.
                            using(StreamWriter sw = File.CreateText(HiddenFilesPath)) {
                                sw.WriteLine(f.FullName);
                            }
                        } else {
                            using(StreamWriter sw = File.AppendText(HiddenFilesPath)) {
                                sw.WriteLine(f.FullName);
                            }
                        }
                    }
                }
            }
        }

        private void ShowFiles() {
            foreach(string path in File.ReadAllLines(HiddenFilesPath)) {
                try {
                    File.SetAttributes(path, FileAttributes.Normal);
                    File.SetAttributes(path, FileAttributes.Normal);
                } catch(Exception e) {
                    WriteToLog("Could not hide file: " + path + " : " + e.Message);
                }
            }
            string[] s = new string[] { string.Empty };
            File.WriteAllLines(HiddenFilesPath, s); // basiclly remove all text from file 
        }

        /// <summary>
        /// Returns all Files with the  " <see cref="_extentionsToHide"/> " extention ending, in all currently opened explorer windows. 
        /// </summary>
        private List<FileInfo> GetFiles() {
            ShellWindows shellWindows = new ShellWindows();
            List<FileInfo> allFiles = new List<FileInfo>();
            foreach(InternetExplorer window in shellWindows) {
                if(window.Document is Shell32.IShellFolderViewDual2 shellWindow) {
                    var currentFolder = shellWindow.Folder.Items().Item();
                    if(currentFolder != null) {
                        DirectoryInfo di = new DirectoryInfo(currentFolder.Path);
                        foreach(string hideFiletype in _extentionsToHide) {
                            allFiles.AddRange(di.GetFiles(hideFiletype));
                        }
                    }
                }
            }
            return allFiles;
        }

        #region Helper Methods
        private void CreateFile(string path) {
            if(File.Exists(path))
                return;
            File.CreateText(path);
        }
        private void DeleteFile(string path) {
            if(!File.Exists(path)) 
                return;
            File.Delete(path);
        }
        private void WriteToLog(string message) {
            using(StreamWriter sw = File.AppendText(LogPath)) {
                sw.WriteLine(DateTime.Now + " : " + message);
            }
        }
        #endregion
    }
}
/** INSTALLATION CODE
 * 1: 
 *      cd C:\Windows\Microsoft.NET\Framework\v4.0.30319
 * 2:
 *      InstallUtil.exe C:\Users\Sven\source\repos\ExplorerFiletypeService\ExplorerFiletypeService\bin\Debug\ExplorerFiletypeService.exe

 * UNINSTALL CODE
 * 1:
 *      InstallUtil.exe -u C:\Users\Sven\source\repos\ExplorerFiletypeService\ExplorerFiletypeService\bin\Debug\ExplorerFiletypeService.exe
 **/