using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using anci.OSTandPSTParser.Inputs;
using anci.OSTandPSTParser.OutputFormats;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace anci.OSTandPSTParser
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        IInputFormat sourceFile;
        private void LoadFile(String filename)
        {
            sourceFile = new InputPST(filename);
            TreeNode node = treeView1.Nodes.Add("ROOT", "Root");
            ParseFolder(sourceFile.RootFolder, node);
            treeView1.ExpandAll();
        }

        private static void ParseFolder(FolderInfo currentFolder, TreeNode node)
        {
            FolderInfoCollection folderInfoCollection = currentFolder.GetSubFolders();

            foreach (FolderInfo folderInfo in folderInfoCollection)
            {
                TreeNode snode = node.Nodes.Add($"{folderInfo.EntryIdString}|{folderInfo.DisplayName}", $"{folderInfo.DisplayName} ({folderInfo.ContentCount})");
                ParseFolder(folderInfo, snode);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == DialogResult.OK) {
                LoadFile(dialog.FileName);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                IOutputFormat outputFile = new OutputPST(dialog.FileName, sourceFile);

                List<TreeNode> selectedNodes = treeView1.GetAllCheckedNodes();

                int i = 0;
                foreach (TreeNode node in selectedNodes)
                {
                    labelState.Text = $"[{++i}/{selectedNodes}] Now exporting \"{node.Text}\"";
                    outputFile.SaveFolder(node);
                }

                outputFile.Close();
            }
            
        }

        private void treeView1_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if (e.Action != TreeViewAction.Unknown)
            {
                if (e.Node.Nodes.Count > 0)
                    e.Node.CheckAllChildNodes(e.Node.Checked);
            }
        }

        private void labelState_Click(object sender, EventArgs e)
        {

        }
    }
}
