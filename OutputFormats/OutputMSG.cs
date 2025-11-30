using anci.OSTandPSTParser.Inputs;
using Aspose.Email;
using Aspose.Email.Mapi;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace anci.OSTandPSTParser.OutputFormats
{
    internal class OutputMSG : IOutputFormat
    {
        public OutputMSG(string folderDest, IInputFormat provider)
        {
            InputProvider = provider;
            Open(DestinationFolder);
        }

        public String DestinationFolder { get; private set; }
        public MsgSaveOptions FileSaveFormat = SaveOptions.DefaultMsgUnicode;

        internal override void Close()
        {

        }

        internal override void Open(string folderName)
        {
            if (!Directory.Exists(folderName))
            {
                Directory.CreateDirectory(folderName);
            }

            DestinationFolder = folderName;
        }

        internal override void SaveFolder(List<String> foldersTree, List<MapiMessage> fileMsg)
        {
            for (int i = 0; i < foldersTree.Count; i++) {
                foldersTree[i] = CheckForInvalidChars(foldersTree[i]);
            }

            //String[] full = [DestinationFolder, ..(foldersTree.ToArray())];
            String newFullPath = Path.Combine(foldersTree.ToArray());
            newFullPath = Path.Combine(DestinationFolder, newFullPath);

            foreach (MapiMessage message in fileMsg) {
                String newFileName = Path.Combine(newFullPath, CheckForInvalidChars(message.Subject));

                message.Save(newFileName, FileSaveFormat);
            }
        }

        private string CheckForInvalidChars(string originalString)
        {
            if (!string.IsNullOrEmpty(originalString))
            {
                return string.Join("_", originalString.Split(Path.GetInvalidFileNameChars()));
            }

            return originalString;
        }
    }
}