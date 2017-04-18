using Microsoft.Office.Interop.OneNote;
using System;
using System.Collections.Generic;
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

            var localNotebook = args[0];
            var remoteNotebook = args[1];

            var app = new Application();
            string localNotebookId;
            app.OpenHierarchy(localNotebook, "", out localNotebookId);

            string xmlLocal;
            app.GetHierarchy(localNotebookId, HierarchyScope.hsSelf, out xmlLocal);

            // get the local notebook name
            XDocument xdoc = XDocument.Parse(xmlLocal);
            var xnameElement = xdoc.Root.Attribute("nickname");

            Console.WriteLine("Opened Notebook " + xnameElement.Value);

            // create the notebook. FYI if there is a notebook with the same name already, this code will just open it
            string remoteNotebookId;
            app.OpenHierarchy(remoteNotebook + xnameElement.Value + " - Remote", string.Empty, out remoteNotebookId, CreateFileType.cftNotebook);

            // just in case there's content, let's just sync everything
            app.SyncHierarchy(remoteNotebookId);

            // we have both open, now let's copy the sections over.  We should do this recursively, since there can be section groups.
            Dictionary<string, string> sectionMappings = new Dictionary<string, string>();
            CopyNotebookRecursively(app, localNotebookId, remoteNotebookId, sectionMappings);

            // A final sync for good luck
            app.SyncHierarchy(remoteNotebookId);

            // worth doing some quick verifications if possible - compare the local and remote page hierarchies and some content.
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
            CopyFolderRecursively(app, oneNs, xdocSourceFolderElement, remoteFolderId, string.Empty);

            // let's sync to make sure things get up to the server
            app.SyncHierarchy(remoteNotebookId);
        }

        private static void CopyFolderRecursively(Application app, XNamespace oneNs, XElement xdocSourceFolderElement, string remoteFolderId, string loggingPrefix)
        {
            // copy each section over
            var sectionElements = xdocSourceFolderElement.Elements(oneNs + "Section");
            foreach (var sectionElement in sectionElements)
            {
                // Copy the section remotely with the same name as the original
                var sectionNameAttribute = sectionElement.Attribute("name");
                var sectionIdAttribute = sectionElement.Attribute("ID");
                string remoteSectionId;

                Console.WriteLine(loggingPrefix + sectionNameAttribute.Value);
                app.OpenHierarchy(sectionNameAttribute.Value + ".one", remoteFolderId, out remoteSectionId, CreateFileType.cftSection);

                app.MergeSections(sectionIdAttribute.Value, remoteSectionId);
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
                CopyFolderRecursively(app, oneNs, sgElement, remoteSgId, loggingPrefix + "    ");
            }
        }
    }
}
