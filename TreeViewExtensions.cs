using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace anci.OSTandPSTParser
{
    public static class TreeViewExtensions
    {
        public static IEnumerable<TreeNode> Descendants(this TreeNodeCollection c)
        {
            foreach (var node in c.OfType<TreeNode>())
            {
                yield return node;

                foreach (var child in node.Nodes.Descendants())
                {
                    yield return child;
                }
            }
        }

        public static List<TreeNode> GetAllCheckedNodes(this TreeView treeview)
        { 
            return treeview.Nodes.Descendants()
                    .Where(n => n.Checked)
                    .ToList();
        }

        public static List<String> GetNamesPath(this TreeNode node) {
            List<String> temp = new List<string>(5);

            do
            {
                temp.Add(node.Name);
                node = node.Parent;
            } while (node != null);

            temp.Reverse();
            return temp;
        }

        public static void CheckAllChildNodes(this TreeNode treeNode, bool nodeChecked)
        {
            foreach (TreeNode node in treeNode.Nodes)
            {
                node.Checked = nodeChecked;
                if (node.Nodes.Count > 0)
                    node.CheckAllChildNodes(nodeChecked);
            }
        }
    }
}
