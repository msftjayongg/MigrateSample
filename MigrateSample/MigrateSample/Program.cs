using Microsoft.Office.Interop.OneNote;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace MigrateSample
{
    /// this sample app shows how to:
    /// 1. Load a local notebook
    /// 2. Make a temporary copy of the local notebook
    /// 3. Create a remote notebook on a SharePoint site
    /// 4. Copy contents of the temporary notebook to the remote notebook
    /// 5. Close and clean up the temporary notebook
    /// 
    /// Recommend this tool is run as administrator. Sometimes get file access denied errors otherwise
    class Program
    {
        // TODO: possible to get this dynamically. e.g. http://stackoverflow.com/questions/934486/how-do-i-get-a-nametable-from-an-xdocument
        private const string strNamespace = "http://schemas.microsoft.com/office/onenote/2013/onenote";
        private XNamespace oneNs = strNamespace;

        static void Main(string[] args)
        {
            var program = new Program();

            program.DoIt(args);
        }

        private void DoIt(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: MigrateSample <localfile> <urlFolderOfRemoteNotebook>");
                Console.WriteLine(@"Example: MigrateSample ""C:\Users\jayongg\Documents\OneNote Notebooks\My Notebook"" ""https://microsoft-my.sharepoint.com/personal/jayongg_microsoft_com/Documents""");
                return;
            }

            string localNotebook = args[0];
            string remoteNotebook = args[1];

            Application app = new Application();

            // Open the local notebook and get its hierarchy
            string localNotebookId;
            app.OpenHierarchy(localNotebook, "", out localNotebookId);
            app.SyncHierarchy(localNotebookId);

            string xmlLocal;
            app.GetHierarchy(localNotebookId, HierarchyScope.hsSections, out xmlLocal);

            // get the local notebook name
            XDocument xdocLocal = XDocument.Parse(xmlLocal);
            XAttribute xnameElement = xdocLocal.Root.Attribute("nickname");
            Console.WriteLine("Opened Notebook " + xnameElement.Value);

            // Create a temporary copy of the local notebook
            string tempLocalNotebook = Path.Combine(System.IO.Path.GetTempPath(), xnameElement.Value);

            // If the temp notebook already exists, delete it
            if (System.IO.Directory.Exists(tempLocalNotebook))
            {
                Console.WriteLine(tempLocalNotebook + " already exists, deleting...");
                System.IO.Directory.Delete(tempLocalNotebook, true /*recursive*/);
            }

            Console.WriteLine("Copying " + localNotebook + " to " + tempLocalNotebook);
            DirectoryCopy(localNotebook, tempLocalNotebook, true /*copySubDirs*/);

            // Open and sync the temporary local notebook
            string tempLocalNotebookId;
            app.OpenHierarchy(tempLocalNotebook, "", out tempLocalNotebookId);
            app.SyncHierarchy(tempLocalNotebookId);

            // Sleep for 3 seconds to give the local notebook time to sync
            // Without this sometimes xmlTempLocal below will be missing content
            System.Threading.Thread.Sleep(3000);

            string xmlTempLocal;
            app.GetHierarchy(tempLocalNotebookId, HierarchyScope.hsSections, out xmlTempLocal);

            // get the temp local notebook name
            XDocument xdocTempLocal = XDocument.Parse(xmlTempLocal);
            xnameElement = xdocTempLocal.Root.Attribute("nickname");
            Console.WriteLine("Opened Temp Notebook " + xnameElement.Value);

            // Get all the first level sections and section groups of the temporary notebook
            var elementsToMove = xdocTempLocal.Root.Elements();

            // create the remote notebook. FYI if there is a notebook with the same name already, this code will just open it
            string remoteNotebookId;
            app.OpenHierarchy(remoteNotebook + xnameElement.Value + " - Remote", string.Empty, out remoteNotebookId, CreateFileType.cftNotebook);

            // just in case there's content, sync the remote notebook
            app.SyncHierarchy(remoteNotebookId);

            // Get the remote hierarchy and add the elements to it
            string xmlRemoteHierarchy;
            app.GetHierarchy(remoteNotebookId, HierarchyScope.hsSections, out xmlRemoteHierarchy);
            XDocument xdocRemote = XDocument.Parse(xmlRemoteHierarchy);

            // If there are existing sectionGroups add the elements before them
            var sectionGroups = xdocRemote.Root.Elements(oneNs + "SectionGroup");
            if (sectionGroups.Count() > 0)
            {
                sectionGroups.First().AddBeforeSelf(elementsToMove);
            }
            // Otherwise just add them under the root
            else
            {
                xdocRemote.Root.Add(elementsToMove);
            }

            // Update the hierarchy with the modified xml
            app.UpdateHierarchy(xdocRemote.ToString());

            // Sync the remote notebook again for good measure
            app.SyncHierarchy(remoteNotebookId);

            // Basic validation
            app.GetHierarchy(remoteNotebookId, HierarchyScope.hsSections, out xmlRemoteHierarchy);
            CompareHierarchy(xmlLocal, xmlRemoteHierarchy);

            // Cleanup - close and delete the temp notebook
            app.CloseNotebook(tempLocalNotebookId);
            if (System.IO.Directory.Exists(tempLocalNotebook))
            {
                Console.WriteLine("Cleaning up " + tempLocalNotebook);
                System.IO.Directory.Delete(tempLocalNotebook, true /*recursive*/);
            }

        }

        private void CompareHierarchy(string xmlLocal, string xmlRemoteHierarchy)
        {
            XDocument xdocLocal = XDocument.Parse(xmlLocal);
            XDocument xdocRemote = XDocument.Parse(xmlRemoteHierarchy);
            int countLocalDescendents = xdocLocal.Descendants().Count<XElement>();
            int countRemoteDescendents = xdocRemote.Descendants().Count<XElement>();
            Console.WriteLine("Validation:");
            Console.WriteLine("Local notebook descendents: " + countLocalDescendents);
            Console.WriteLine("Remote notebook descendents: " + countRemoteDescendents);
        }

        private static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            // Get the subdirectories for the specified directory.
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);

            if (!dir.Exists)
            {
                Console.WriteLine("Source directory does not exist or could not be found: " + sourceDirName);
                return;
            }

            DirectoryInfo[] dirs = dir.GetDirectories();
            // If the destination directory doesn't exist, create it.
            if (!Directory.Exists(destDirName))
            {
                Directory.CreateDirectory(destDirName);
            }

            // Get the files in the directory and copy them to the new location.
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string temppath = Path.Combine(destDirName, file.Name);
                file.CopyTo(temppath, false);
            }

            // If copying subdirectories, copy them and their contents to new location.
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    string temppath = Path.Combine(destDirName, subdir.Name);
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs);
                }
            }
        }
    }
}
