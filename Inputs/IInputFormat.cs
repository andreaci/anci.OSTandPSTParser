using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace anci.OSTandPSTParser.Inputs
{
    internal abstract class IInputFormat
    {
        internal abstract FolderInfo RootFolder { get; }

        internal abstract FolderInfo GetFolderById(string folderId);

        internal abstract MapiMessage ExtractMessage(MessageInfo message);
    }
}
