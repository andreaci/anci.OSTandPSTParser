using anci.OSTandPSTParser.Inputs;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace anci.OSTandPSTParser.OutputFormats
{
    internal abstract class IOutputFormat
    {
        public IInputFormat InputProvider { get; set; }

        internal abstract void Open(string filename);


        internal abstract void SaveFolder(List<String> foldersTree, List<MapiMessage> fileMsg);
        internal abstract void Close();


        internal void SaveFolder(TreeNode node) {

            List<MapiMessage> messagesList = GetMesssagesFromFolder(node, out List<string> nodePath);
            SaveFolder(nodePath, messagesList);
        }
        protected List<MapiMessage> GetMesssagesFromFolder(TreeNode node, out List<string> nodePath)
        {
            nodePath = node.GetNamesPath();
            nodePath = nodePath.Skip(1).Select(n => n.Split('|')[1]).ToList();

            FolderInfo folder = InputProvider.GetFolderById(node.Name.Split('|')[0]);

            MessageInfoCollection messages = folder.GetContents();

            List<MapiMessage> messagesList = new List<MapiMessage>(messages.Count);
            foreach (MessageInfo message in messages)
            {
                MapiMessage mapiTemp = InputProvider.ExtractMessage(message);
                messagesList.Add(mapiTemp);
            }

            return messagesList;
        }
    }
}
