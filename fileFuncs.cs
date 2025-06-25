using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace PowerRevUp
{
    internal class fileFuncs
    {
        /// <summary>
        /// CreateContext checks if we need to create an entry in the context menu and does it if so
        /// </summary>
        /// <returns>True if operation was successful, false otherwise</returns>
        public static bool CreateContext()
        {
            //Check if we need to createContext and if so elevate priviledges
            //Read the registry and check for the regedit
            try
            {
               
                RegistryKey rk = Registry.ClassesRoot.OpenSubKey("*", true).OpenSubKey("shell", true);

                //Setup access security to allow the registry to be edited
                RegistrySecurity rs = new RegistrySecurity();
                rs.AddAccessRule(new RegistryAccessRule("admin", RegistryRights.FullControl, AccessControlType.Allow));
                rk.SetAccessControl(rs);


                string[] keyNames = rk.GetSubKeyNames();
                foreach (string key in keyNames)
                {
                    //If there is already a GPADocIssue key in registry, there is already an entry in the context menu
                    if (key == "PowerRevUp")
                    {
                        rk.Close();
                        return false;
                    }
                }
                
                try
                {

                    AdminRelauncher();
                    //Creates entry
                    RegistryKey parentKey = rk.CreateSubKey("PowerRevUp", true);
                    //Allows up to 100 files to be passed in
                    parentKey.SetValue("MultiSelectModel", "Player");

                    ////Sets the icon to the rubber duck
                    parentKey.SetValue("Icon", "\"C:\\GPA Apps\\PowerRevUp\\faviconArrow.ico\"");

                    //Sets the value of the location of the exe - probably need to change this to find the exe location before release.
                    RegistryKey childKey = parentKey.CreateSubKey("command", true);
                    childKey.SetValue(null, "\"C:\\GPA Apps\\PowerRevUp\\PowerRevUp.exe\" \"%1\" ");
                    rk.Close();

                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show($"{ex.Message} Error occurred while creating shortcut.");
                    rk.Close();
                }
                return true;
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Error adding PowerRevUp to context menu.");
                return false;
            }
        }

        /// <summary>
        /// AdminRelauncher looks in the registry for a key which contains a context menu shortcut to PowerRevUp. If its already there it stops, otherwise it then checks if DocIssue was launched as admin and relaunches if not.
        /// </summary>
        /// <returns>False if not needed to relaunch as admin. True otherwise</returns>
        public static bool AdminRelauncher()
        {
            RegistryKey regKey = Registry.ClassesRoot.OpenSubKey("*", true).OpenSubKey("shell", true);
            string[] keyNames = regKey.GetSubKeyNames();
            foreach (string key in keyNames)
            {
                if (key == "PowerRevUp")
                {
                    regKey.Close();
                    return true;
                    //Don't relaunch because the shortcut already exists
                }
            }

            if (!IsRunAsAdmin())
            {
                ProcessStartInfo proc = new ProcessStartInfo();
                proc.UseShellExecute = true;
                proc.WorkingDirectory = Environment.CurrentDirectory;
                proc.FileName = System.Windows.Forms.Application.ExecutablePath;

                proc.Verb = "runas";
                try
                {
                    Process.Start(proc);
                    Environment.Exit(0);
                    return true;
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                    return false;
                }
            }
            return false;
        }

        /// <summary>
        /// IsRunAsAdmin checks if the current app has been run as an administrator and returns true if so.
        /// </summary>
        /// <returns>True if app run as admin, false otherwise.</returns>
        public static bool IsRunAsAdmin()
        {
            try
            {
                WindowsIdentity identity = WindowsIdentity.GetCurrent();
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        /// <summary>
        /// Checks the base dir for an args folder, reads all of the txt files in it and then returns a list of all command line arguments from other instances
        /// </summary>
        /// <returns>A list of all the args collected from the written files</returns>
        public static List<string> GetArguments()
        {
            string argsDir = AppDomain.CurrentDomain.BaseDirectory + "args";
            //check if there are any files in the dir that don't start with args
            string[] files = Directory.GetFiles(argsDir);
            //Give the first instance this mutex to prevent other threads from trying to read/delete the files that the main thread is trying to access
          
            //CheckArgs(files);

            //I don't think that any of the instances are getting this mutex 

            List<string> argsList = new List<string>();
            string[] _files = Directory.GetFiles(argsDir);
           
            string line;
            foreach (string fileName in _files)
            {
                using (StreamReader reader = File.OpenText(fileName))
                {
                    //Keep reading till EOF
                    while ((line = reader.ReadLine()) != null)
                    {
                        //If the extension isnt dll and the file isnt a folder - would need to change this to allow dll's to be passed in. Its like this right now because otherwise the PowerRevUp.dll gets passed in.
                        if ((System.IO.Path.GetExtension(line) != (".dll")))
                        {
                            //add file name to args
                            argsList.Add(line);
                        }
                    }
                }
            }
            // string[] argsArray = argsList.ToArray<string>();
            //Delete all of the files
            
            return argsList;
        }

        public static void CheckArgs(string[] files)
        {
            try
            {
                foreach (string file in files)
                {
                    //if the filename doesnt start with args
                    if (!System.IO.Path.GetFileName(file).StartsWith("args"))
                    {

                        File.Delete(file);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error checking files for arguments");
            }

        }

        public static void GetSingleInstance()
        {
            bool isNewInstance = true;
            Guid guid = Guid.NewGuid();

            string[] newArgs = Environment.GetCommandLineArgs();
            //need to do a check here to make sure that the folder with doc issue in it has a folder called args to store the arguments
            string fileName = "args " + Guid.NewGuid().ToString() + ".txt";
            string path = AppDomain.CurrentDomain.BaseDirectory + "args\\" + fileName;
            Mutex m_mutex = new Mutex(false, "be6d5f08ee9442d488ceebe32fecb1a9", out isNewInstance);
            //Write each to a seperate file with the name "args + GUID" and then loop through all of the files in that folder and see which are args + guid and then compile all of those into a list

            using (StreamWriter writer = new StreamWriter(path))
            {
                Parallel.ForEach(newArgs, arg =>
                 {
                     writer.WriteLine(arg);
                 });
                writer.Close();
            }

            if (!isNewInstance && Process.GetProcessesByName("PowerRevUp").Length > 1)
            {
                Environment.Exit(0);
            }

        }

        public static void FileDeleter(string[] files)
        {
            try
            {
                foreach (string file in files)
                {
                    File.Delete(file);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Finds the revision value of a file and increases it by one. 
        /// </summary>
        /// <param name="file2"></param>
        /// <returns>String location of the rev'd up file.</returns>
        public static string RevFileUp(string file2)
        {
            bool isMatch = false;
            //Regex match for the rev num/letter
            Regex r = new Regex(@"(?<= r).+(?=\..+$)");
            Match m;
            int matchLength = 1;
            string revNumber = String.Empty;
            string revLetter = String.Empty;
            string mString = String.Empty;
            try
            {
                m = r.Match(file2);
                mString = m.ToString();
                
                if (m.Success)
                {
                    isMatch = true;
                    if (m.ToString().Length > 1)
                    {
                        matchLength = m.ToString().Length;
                    }
                    //Need to then find if the length is greater than one, and change the functionality if it is
                
                    //Need to add handling for number/alphanumeric revisions
                    //Check if the match is a number on its own
                    bool isNumber = int.TryParse(m.ToString(), out int n);
                    if(isNumber)
                    {
                        revNumber = n.ToString();
                        isMatch = false;
                    }

                    //Check if the match is alphanumeric
                    Regex letterNumReg = new Regex(@"(?'num'\d*)(?'letter'\D*)");
                    Match letterNumMatch = letterNumReg.Match(mString);

                    if (letterNumMatch.Groups["num"].Value != null && letterNumMatch.Groups["letter"].Value != null)
                    {
                        //Is alphanumeric
                        revNumber = letterNumMatch.Groups["num"].Value;
                        mString = letterNumMatch.Groups["letter"].Value;
                    }
                    
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            m = r.Match(file2);
            

            
            if (isMatch)
            {
                int colNum = GetExcelColNumber(mString) + 1;
                revLetter = GetExcelColName(colNum, matchLength);
            }
            else
            {
                revLetter = "A";
            }

            //FINAL STRING TO RENAME THE FILE 
            string rename;

            string rev = $"r{revNumber + revLetter}";
            string revExtension = rev + Path.GetExtension(file2);
            string dirName = Path.GetDirectoryName(file2);

            string fileName = Path.GetFileNameWithoutExtension(file2);
            //Need to find the current revision and remove it from the end of the string before continuing
            Regex removeRev = new Regex(@"\s[r].+(?=)*$");
            Match _m = removeRev.Match(fileName);
            if (_m.Success)
            {
                fileName = Regex.Replace(fileName, @"\s[r].+(?=)*$", "");
            }
            else
            {
                //Give us a yell
                //MessageBox.Show($"Error in renaming file. Exiting...");
                //Environment.Exit(0);
            }



            string filePathNoExtension = Path.Combine(dirName, fileName);
            rename = filePathNoExtension + " " + revExtension;

            //Renames the file to have a revision after it 
            try
            {
                File.Move(file2, rename, true);
                return rename;
                //System.Windows.MessageBox.Show($"File2: {file2}, rename: {rename}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error in moving: \n{ex.Message}");
                return "";

            }
        }

        /// <summary>
        /// Takes a number representing a letter value and turns it into said letter value. Ie. A =1, B = 2, AA = 27 etc.
        /// </summary>
        /// <param name="colNumber">Number representing a string in a format like excel columns</param>
        /// <returns>Conversion of an int representing a letter val to a string</returns>
        public static string GetExcelColName(int colNumber, int matchLength)
        {

            //Takes the rev and increases it but allows Z => AA, ZZ => AAA etc.
            int dividend = colNumber;
            string colName = String.Empty;

            int modulo;
            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                colName = Convert.ToChar(65 + modulo).ToString() + colName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return colName;
        }

        /// <summary>
        /// Takes a string and turns it into a number <br></br> For example, A = 1, B = 2 ... AA = 27 ... AAA = 703
        /// </summary>
        /// <param name="colName">Current letter of the revision</param>
        /// <returns>A number representation of a letter. (ie. A = 1, AA = 27, AAA = 703 etc.)</returns>
        public static int GetExcelColNumber(string colName)
        {
            //Takes a letter and turns it into a number so that it can be incremented
            char[] chars = colName.ToUpperInvariant().ToCharArray();
            int sum = 0;
            for (int i = 0; i < chars.Length; i++)
            {
                sum *= 26;
                sum += (chars[i] - 'A' + 1);
            }
            //MessageBox.Show(sum.ToString());
            return sum;
        }

        /// <summary>
        /// Cleans up the args folder by deleting all of the files inside of it.
        /// </summary>
        public static void CleanupFiles()
        {
            string argsDir = AppDomain.CurrentDomain.BaseDirectory + "args";
            string[] _files = Directory.GetFiles(argsDir);
            //Gonna need a cleanup function to get rid of the folder files if they are still there once the app has closed
            try
            {
                foreach (string file in _files)
                {
                    File.Delete(file);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Shows a warning if a shortcut is passed in. Then cleans up the args folder before exiting the program.
        /// </summary>
        public static void ShortcutWarning()
        {
            DialogResult res = MessageBox.Show("The only file passed in is a shortcut which breaks some functionality of PowerRevUp, please try again with the original file. \n\nIf you selected multiple files and this error occurred, try right clicking on another file that isn't a shortcut.", "Error!", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            CleanupFiles();
            Environment.Exit(0);
        }
    }
}
