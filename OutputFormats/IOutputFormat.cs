using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace anci.OSTandPSTParser.OutputFormats
{
    internal abstract class IOutputFormat
    {
        internal abstract void Open(string filename);

        internal abstract void SaveFolder(TreeNode node);
        internal abstract void Close();
    }
}
