using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using anci.OSTandPSTParser.Inputs;
using System;
using System.Collections.Generic;
using System.IO;

namespace anci.OSTandPSTParser.OutputFormats
{
    internal class OutputPST : IOutputFormat
    {
        public OutputPST(string fileName, IInputFormat provider)
        {
            InputProvider = provider;
            Open(fileName);
        }

        private PersonalStorage DataFile { get; set; }


        internal override void Open(string filename)
        {
            if (File.Exists(filename))
                DataFile = PersonalStorage.FromFile(filename);
            else
                DataFile = PersonalStorage.Create(filename, FileFormatVersion.Unicode);
        }


        internal override void SaveFolder(List<String> foldersTree, List<MapiMessage> fileMsg)
        {
            FolderInfo inboxFolder = DataFile.RootFolder;
            foreach (String folder in foldersTree)
            {
                FolderInfo temp = inboxFolder.GetSubFolder(folder);
                if (temp != null)
                    inboxFolder = temp;
                else
                {
                    inboxFolder.AddSubFolder(folder);
                    inboxFolder = inboxFolder.GetSubFolder(folder);
                }
            }

            if (inboxFolder != null)
            {
                foreach (MapiMessage message in fileMsg)
                    inboxFolder.AddMessage(message);
            }
        }

        internal override void Close()
        {
            DataFile.Dispose();
        }
    }
}
