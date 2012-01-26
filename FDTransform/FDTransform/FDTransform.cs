using System;
using System.IO;
using System.IO.Packaging;
using System.Xml.Xsl;
using System.Security.Permissions;
using Microsoft.SharePoint;
using System.Xml;
using System.Xml.XPath;

namespace FDTransform.EventReceivers
{
    public class FDTransformEventReceiver : SPItemEventReceiver
    {
        // TODO: Need a better named list? Maybe a name that signifies internal?
        // We keep the transform file in SharePoint in this following Document List
        string TransformsListName = "Faculty Data XSLT to HTML";
        // under this item name.
        string TransformItemName = "FDR.xslt";

        // We place the result in a new item with the same name.
        string ResultsListName = "HTMLDocs";

        // This function constructs the name for the result from 
        // the original name. The given name is an InfoPath form
        // and has an extension of xml.
        protected string getResultsItemNameFromName(string name)
        {
            string newName = name;
            if (newName.EndsWith(".xml"))
            {
                newName = name.Substring(0, name.Length - 4);
            }
            return newName + ".htm";
        }
            
        protected Stream GetTransformStream(SPWeb web)
        {
            SPFile transformF = web.GetFile("FDR/Faculty Data XSLT to HTML/FDR.xslt");
            // This doesn't work:;  transformS = transformF.OpenBinaryStream();
            SPFolder folder = web.GetFolder(TransformsListName);
            foreach (SPFile file in folder.Files)
            {
                if (file.Exists)
                {
                    if (file.Name == TransformItemName)
                    {
                        return file.OpenBinaryStream();
                    }
                }
            }
            // Perhaps raise an exception here.
            return null;
        }

        private void doit(SPItemEventProperties properties)
        {
            // This is our Context. We get files from it.
            SPWeb oWeb = properties.OpenWeb();

            // Get the transform.
            Stream transformS = GetTransformStream(oWeb);
            XslCompiledTransform trans = new XslCompiledTransform();
            trans.Load(new XmlTextReader(transformS));

            // Create a Temp file to write the result to.
            // TODO: Probably could do this in memory with byte[] as they may
            // not be that large.
            string tempDoc = Path.GetTempFileName();
            new FileInfo(tempDoc).Attributes = FileAttributes.Temporary;
            FileStream partStream = File.Open(tempDoc, FileMode.Create, FileAccess.Write);

            // Get the selected document.
            SPFile xml = properties.ListItem.File;
            XPathDocument doc = new XPathDocument(xml.OpenBinaryStream());

            // Transform the selected document and close.
            trans.Transform(doc, null, partStream);
            partStream.Close();

            // Store the result in the Results Folder
            SPFolder resultsfolder = oWeb.GetFolder(ResultsListName);
            string filename = getResultsItemNameFromName(xml.Name);

            FileStream inStream = File.Open(tempDoc, FileMode.Open, FileAccess.Read);

            // We will overwrite the item if it's updating.
            resultsfolder.Files.Add(filename, inStream);

            // Clean up, delete the temp file.
            File.Delete(tempDoc);
        }

       public override void ItemAdded(SPItemEventProperties properties)
       {
           base.ItemAdded(properties);
           doit(properties);
       }

       /// <summary>
       /// An item is being updated.
       /// </summary>
       public override void ItemUpdated(SPItemEventProperties properties)
       {
           base.ItemUpdated(properties);
           doit(properties);
       }


    }
}
