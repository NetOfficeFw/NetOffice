using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor
{
    public static class TreeViewExtensions
    {
        public static Utils.Registry.UtilsRegistry Registry(this TreeNode node)
        {
            return (Utils.Registry.UtilsRegistry)node.Tag;
        }

        public static Utils.Registry.UtilsRegistryKey RegistryKey(this TreeNode node)
        {
            return (Utils.Registry.UtilsRegistryKey)node.Tag;
        }

        public static TreeNode Root(this TreeNode node)
        {
            TreeNode result = node;
            while (null != result.Parent)
            {
                result = result.Parent;
            }
            return result;
        }
    }
}
