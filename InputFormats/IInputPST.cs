using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace anci.OSTandPSTParser.Inputs
{
    internal class InputPST : IInputFormat
    {
        PersonalStorage sourceFile;

        internal InputPST(String filename) {

            sourceFile = PersonalStorage.FromFile(filename);
        }

        internal override FolderInfo RootFolder => sourceFile.RootFolder;

        internal override MapiMessage ExtractMessage(MessageInfo message) => sourceFile.ExtractMessage(message);

        internal override FolderInfo GetFolderById(string folderId) => sourceFile.GetFolderById(folderId);
    }
}
