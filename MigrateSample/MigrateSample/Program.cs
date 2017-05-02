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
    /// 2. Load a remote notebook
    /// 3. copy contents of the local notebook to the remote notebook
    /// 4. close the local notebook
    /// 
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

            string tempLocalNotebookId;
            app.OpenHierarchy(tempLocalNotebook, "", out tempLocalNotebookId);
            app.SyncHierarchy(tempLocalNotebookId);
            System.Threading.Thread.Sleep(1000);

            string xmlTempLocal;
            app.GetHierarchy(tempLocalNotebookId, HierarchyScope.hsSections, out xmlTempLocal);

            // get the local notebook name
            XDocument xdocTempLocal = XDocument.Parse(xmlTempLocal);
            xnameElement = xdocTempLocal.Root.Attribute("nickname");
            Console.WriteLine("Opened Temp Notebook " + xnameElement.Value);


            // Get all the elements of the temporary notebook
            var elementsToMove = xdocTempLocal.Root.Descendants();

            // Full XDocument version
            // Get the full hierarchy
            //string xmlFullHierarchy;
            //app.GetHierarchy("", HierarchyScope.hsSections, out xmlFullHierarchy);
            //XDocument xdocFull = XDocument.Parse(xmlFullHierarchy);
            //var notebookElements = xdocFull.Descendants(oneNs + "Notebook");

            //foreach (var notebookElement in notebookElements)
            //{
            //    if (notebookElement.Attribute("ID").Value == remoteNotebookId)
            //    {
            //        notebookElement.Add(elementsToMove);
            //    }
            //}

            // create the remote notebook. FYI if there is a notebook with the same name already, this code will just open it
            string remoteNotebookId;
            app.OpenHierarchy(remoteNotebook + xnameElement.Value + " - Remote", string.Empty, out remoteNotebookId, CreateFileType.cftNotebook);

            // just in case there's content, sync the remote notebook
            app.SyncHierarchy(remoteNotebookId);

            // Just remote XDocument version
            string xmlRemoteHierarchy;
            app.GetHierarchy(remoteNotebookId, HierarchyScope.hsSections, out xmlRemoteHierarchy);
            XDocument xdocRemote = XDocument.Parse(xmlRemoteHierarchy);
            xdocRemote.Root.Add(elementsToMove);
            //XmlDocument xmlDocUpdated = new XmlDocument();
            //xmlDocUpdated.LoadXml(xdocRemote.ToString());

            // Update the hierarchy with the modified xml
            app.UpdateHierarchy(xdocRemote.ToString());

            // Sync the remote notebook again for good measure
            app.SyncHierarchy(remoteNotebookId);

            // worth doing some quick verifications if possible - compare the local and remote page hierarchies and some content.

            // Cleanup - close and delete the temp notebook
            app.CloseNotebook(tempLocalNotebookId);
            if (System.IO.Directory.Exists(tempLocalNotebook))
            {
                Console.WriteLine("Cleaning up " + tempLocalNotebook);
                System.IO.Directory.Delete(tempLocalNotebook, true /*recursive*/);
            }

        }

        private void CopyNotebookRecursively(Application app, string localNotebookId, string remoteNotebookId, Dictionary<string, string> sectionMappings)
        {
            // get the hierarchy
            string xmlLocal;
            app.GetHierarchy(localNotebookId, HierarchyScope.hsSections, out xmlLocal);

            string xmlRemote;
            app.GetHierarchy(remoteNotebookId, HierarchyScope.hsSelf, out xmlRemote);

            XDocument xdocLocal = XDocument.Parse(xmlLocal);

            var remoteFolderId = remoteNotebookId;
            var xdocSourceFolderElement = xdocLocal.Root;
            CopyFolderRecursively(app, oneNs, xdocSourceFolderElement, localNotebookId, remoteFolderId, string.Empty);

            // let's sync to make sure things get up to the server
            app.SyncHierarchy(remoteNotebookId);
        }

        private static void CopyFolderRecursively(Application app, XNamespace oneNs, XElement xdocSourceFolderElement, string localNotebookId, string remoteFolderId, string loggingPrefix)
        {
            // copy each section over
            var sectionElements = xdocSourceFolderElement.Elements(oneNs + "Section");
            foreach (var sectionElement in sectionElements)
            {
                // Copy the section remotely with the same name as the original
                var sectionNameAttribute = sectionElement.Attribute("name");
                var sectionIdAttribute = sectionElement.Attribute("ID");

                string strDotOne = sectionNameAttribute.Value + ".one";
                string tempDotOnePath = Path.Combine(System.IO.Path.GetTempPath(), strDotOne);

                // Publish the section to a temporary .one file
                Console.WriteLine(loggingPrefix + "Publishing " + sectionNameAttribute.Value + " ID: " + sectionIdAttribute.Value);

                // delete any preexisting .one file
                File.Delete(tempDotOnePath);

                // publish the section, given the section ID.
                app.Publish(sectionIdAttribute.Value, tempDotOnePath);

                // Open the temporary .one
                string tempLocalId;
                app.OpenHierarchy(tempDotOnePath, "", out tempLocalId);

                // Get its xml
                XmlDocument xmlDoc = new XmlDocument();
                string hierarchy = "";
                app.GetHierarchy(tempLocalId, HierarchyScope.hsSelf, out hierarchy);
                xmlDoc.LoadXml(hierarchy);

                // Move the temporary .one to the remote notebook
                app.UpdateHierarchy(xmlDoc.OuterXml);
            }

            // now lets do this for each section group
            var sgElements = xdocSourceFolderElement.Elements(oneNs + "SectionGroup");
            foreach (var sgElement in sgElements)
            {
                // skip recycle bin
                var isRecycleAttribute = sgElement.Attribute("isRecycleBin")?.Value;
                if (isRecycleAttribute != null && isRecycleAttribute == "true")
                    continue;

                // Create the section group remotely, with name of source section group
                var sgNameAttribute = sgElement.Attribute("name");
                var sgIdAttribute = sgElement.Attribute("ID");
                string remoteSgId;
                Console.WriteLine(loggingPrefix + sgNameAttribute.Value);

                app.OpenHierarchy(sgNameAttribute.Value, remoteFolderId, out remoteSgId, CreateFileType.cftFolder);

                // Copy the contents of the source section group to remote section group
                CopyFolderRecursively(app, oneNs, sgElement, localNotebookId, remoteSgId, loggingPrefix + "    ");
            }
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
