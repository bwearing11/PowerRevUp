using System.Diagnostics;
using System.Windows.Forms;

namespace PowerRevUp
{
    internal class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            fileFuncs.GetSingleInstance();
            Thread.Sleep(100);
            
            //BY HERE NONE OF THE EXTRA INSTANCES ARE STILL OPEN

            //Check if the context menu add in has been made, if not make it
            bool isContext = fileFuncs.CreateContext();      
           if(isContext)
            {
                MessageBox.Show("PowerRevUp added to the right click menu. ", "PowerRevUp Context Menu", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            List<string> revdFiles = new List<string>();
            //Not-IPC to get all of the selected file names from the context menu
            //string[] tempArgs = Environment.GetCommandLineArgs();
            string dirString = AppDomain.CurrentDomain.BaseDirectory + "args";
            int fileCount = Directory.GetFiles(dirString).Length;
            //Very important sleep
            Thread.Sleep(20 * fileCount);
            List<string> args = fileFuncs.GetArguments();

            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            folderBrowserDialog1.InitialDirectory = Directory.GetParent(args[0]).ToString();
            folderBrowserDialog1.ShowNewFolderButton = true;
            DialogResult result = folderBrowserDialog1.ShowDialog();
            string selectedPath = "";
            bool isPath = false;                                        //Whether or not they chose a folder to move to, if they didn't then just rename and don't move at all
            if (result == DialogResult.OK)
            {
                selectedPath = folderBrowserDialog1.SelectedPath;
                isPath = true;
            }

            if (!isPath)
            {
                //If a folder wasn't selected do it as normal
                foreach (string arg in args)
                {
                    //Need to do a shortcut check first 
                    if (Path.GetExtension(arg) == ".exe" && args.Count <= 2)
                    {
                        fileFuncs.ShortcutWarning();
                    }
                    string revFile = fileFuncs.RevFileUp(arg);
                    revdFiles.Add(revFile);
                }
            }
            else
            {
                
                //If a folder was selected, copy the file to a temp directory, rev it up and then move it to the appropriate dir.
                foreach (string arg in args)
                {
                    //Need to do a shortcut check first 
                    if (Path.GetExtension(arg) == ".exe" && args.Count <= 2)
                    {
                        fileFuncs.ShortcutWarning();
                    }

                    //Copy the file to a temp dir
                    string tempDir = Path.GetTempPath();
                    string tempPath = Path.Combine(tempDir, Path.GetFileName(arg));
                    try
                    {
                        File.Copy(arg, tempPath);
                    }catch(Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    //Then rev it up  
                    string revD = fileFuncs.RevFileUp(tempPath);
                    
                    
                    string argDir = Directory.GetParent(arg).ToString();
                    string argPath = Path.Combine(argDir, Path.GetFileName(revD));
                    //Then move it to the right spot
                    try
                    {
                        File.Move(revD, argPath, false);
                        revdFiles.Add(argPath);
                        
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show($"Error moving file from temp directory to original folder. {ex.Message}");
                    }
                }
            }
                      
            

            if (isPath)
            {
                string movePath;
                //Move the files before exiting
                foreach (string file in revdFiles)
                {
                    movePath = Path.Combine(selectedPath, Path.GetFileName(file));
                    try {
                        File.Move(file, movePath, false);
                    }catch(Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    }
            }
            fileFuncs.CleanupFiles();

            //Action successfully completed messagebox
            string final = $"Successfully Rev'd Up {args.Count} File(s)";
            string finalString = final;
            string dirInfo = String.Empty;

            if (isPath)
            {
                dirInfo = $" And Moved To {Path.GetFileName(selectedPath)}";
                finalString += dirInfo;
            }

            MessageBox.Show(finalString, "PowerRevUp", MessageBoxButtons.OK, MessageBoxIcon.Information);
            
        }
    }
}