using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using NetOffice.DeveloperToolbox.Utils.Registry;

namespace NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor
{
    public static class UtilsRegistryKeyExtensions
    {
        public static UtilsRegistryKey Next(this UtilsRegistryKey node)
        {
            if (node.IsRoot())
                return node.Keys.FirstOrDefault();

            if (node.HasChildren())
                return node.Keys.FirstOrDefault();

            if (node.IsLastChildren())
                return node.NextSibling();
            else
                return node.Parent.Keys[node.IndexOf() + 1];
        }

        private static UtilsRegistryKey NextSibling(this UtilsRegistryKey node)
        {
            node = node.Parent;
            while (null != node)
            {
                if (!node.IsLastChildren())
                    return null != node.Parent ? node.Parent.Keys[node.IndexOf() + 1] : null;
                node = node.Parent;
            }
            return null;
        }

        private static bool HasChildren(this UtilsRegistryKey node)
        {
            return node.Keys.Count > 0;
        }

        private static bool IsLastChildren(this UtilsRegistryKey node)
        {
            return null != node.Parent ? node.Parent.Keys.LastName == node.Name : false;
        }

        private static bool IsRoot(this UtilsRegistryKey node)
        {
            return null == node.Parent;
        }

        private static int IndexOf(this UtilsRegistryKey node)
        {
            return node.Keys.IndexOf(node.Name);
        }
    }
}