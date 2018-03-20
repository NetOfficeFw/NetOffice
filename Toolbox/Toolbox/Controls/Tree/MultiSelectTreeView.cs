using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.Controls.Tree
{
    /// <summary>
    /// taken from: http://www.mycsharp.de/wbb2/thread.php?threadid=82474
    /// </summary>
    public class MultiSelectTreeView : TreeView
    {
        private List<TreeNode> selected_nodes = null;
        private TreeNode last_node = null;
        private bool only_focus = false;

        public List<TreeNode> SelectedNodes
        {
            get
            {
                return selected_nodes;
            }
            set
            {
                SelectNodes(value);
            }
        }

        public MultiSelectTreeView()
        {
            base.HideSelection = false;
            base.DrawMode = TreeViewDrawMode.OwnerDrawText;
            selected_nodes = new List<TreeNode>();
        }

        protected override void OnBeforeSelect(TreeViewCancelEventArgs e)
        {
            if (!only_focus)
            {
                base.OnBeforeSelect(e);
                NewSelection(e.Node);
            }
        }

        protected override void OnAfterSelect(TreeViewEventArgs e)
        {
            if (!only_focus)
            {
                base.OnAfterSelect(e);
            }
        }

        protected override void OnDrawNode(DrawTreeNodeEventArgs e)
        {
            if (e.Node.IsVisible)
            {
                Color nBackColor = this.BackColor;
                Color nForeColor = this.ForeColor;
                if (selected_nodes.Contains(e.Node))
                {
                    nBackColor = SystemColors.Highlight;
                    nForeColor = SystemColors.HighlightText;
                }
                // Retrieve the node font. If the node font has not been set,
                // use the TreeView font.
                Font nFont = e.Node.NodeFont;
                if (nFont == null)
                    nFont = base.Font;

                Rectangle nBounds = e.Node.Bounds;
                int nTextX = nBounds.X;
                int nTextY = nBounds.Y + (nBounds.Height - nFont.Height) / 2;

                e.Graphics.FillRectangle(new SolidBrush(nBackColor), nBounds);
                e.Graphics.DrawString(e.Node.Text, nFont, new SolidBrush(nForeColor), nTextX, nTextY);
                // If the node has focus, draw the focus rectangle large, making
                // it large enough to include the text of the node tag, if present.
                if ((e.State & TreeNodeStates.Focused) != 0)
                {
                    using (Pen focusPen = new Pen(Color.Black))
                    {
                        focusPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dot;
                        Rectangle focusBounds = nBounds;
                        focusBounds.Size = new Size(focusBounds.Width - 1,
                        focusBounds.Height - 1);
                        e.Graphics.DrawRectangle(focusPen, focusBounds);
                    }
                }
            }
        }

        private void RedrawNode(TreeNode node)
        {
            Invalidate(node.Bounds);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            only_focus = false;
            TreeNode node = GetNodeAt(e.X, e.Y);
            if (node != null)
            {
                if (node.Bounds.Contains(e.X, e.Y))
                {
                    if (base.SelectedNode != node)
                        base.SelectedNode = node;
                }
            }
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            only_focus = e.Control && (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down ||
                                       e.KeyCode == Keys.Left || e.KeyCode == Keys.Right);
            if (only_focus)
            {
                e.SuppressKeyPress = true;
                switch (e.KeyCode)
                {
                    case Keys.Up:
                        if (base.SelectedNode.PrevVisibleNode != null)
                            base.SelectedNode = base.SelectedNode.PrevVisibleNode;
                        break;
                    case Keys.Down:
                        if (base.SelectedNode.NextVisibleNode != null)
                            base.SelectedNode = base.SelectedNode.NextVisibleNode;
                        break;
                    case Keys.Left:
                        base.SelectedNode.Collapse(true);
                        break;
                    case Keys.Right:
                        base.SelectedNode.Expand();
                        break;
                }
            }
            else if (e.KeyCode == Keys.Space)
            {
                TreeNode node = base.SelectedNode;
                if (e.Control || !selected_nodes.Contains(node))
                {
                    base.SelectedNode = null;
                    base.SelectedNode = node;
                }
            }
        }

        protected override void OnBeforeCollapse(TreeViewCancelEventArgs e)
        {
            if (RemoveSelectionBeforeCollapse(e.Node))
            {
                SelectedNode = e.Node;
            }
            base.OnBeforeCollapse(e);
        }

        private bool RemoveSelectionBeforeCollapse(TreeNode node)
        {
            bool a = selected_nodes.Contains(node);
            if (a)
            {
                selected_nodes.Remove(node);
                RedrawNode(node);
            }
            foreach (TreeNode n in node.Nodes)
            {
                a = a || RemoveSelectionBeforeCollapse(n);
            }
            return a;
        }

        private void NewSelection(TreeNode node)
        {
            bool bControl = (ModifierKeys == Keys.Control);
            bool bShift = (ModifierKeys == Keys.Shift);
            if (!bShift) last_node = node;
            if (bControl)
            {
                ToogleNodeSelection(node);
            }
            else if (bShift)
            {
                SelectBetweenNodes(last_node, node);
            }
            else
            {
                SelectSingleNode(node);
            }
        }

        private void SelectNode(TreeNode node)
        {
            int currentNodeLevel = node.Level;
            foreach (TreeNode item in selected_nodes)
            {
                if (item.Level != currentNodeLevel)
                    return;
            }

            if (!selected_nodes.Contains(node))
            {
                selected_nodes.Add(node);
                RedrawNode(node);
                SortSelectedNodes();
            }
        }

        private void DeselectNode(TreeNode node)
        {
            if (selected_nodes.Contains(node))
            {
                selected_nodes.Remove(node);
                RedrawNode(node);
            }
        }

        private void ToogleNodeSelection(TreeNode node)
        {
            if (selected_nodes.Contains(node))
                DeselectNode(node);
            else
                SelectNode(node);
        }

        private void SelectSingleNode(TreeNode node)
        {
            SelectNodes(new List<TreeNode>(new TreeNode[] { node }));
        }

        private void SelectNodes(List<TreeNode> nodes)
        {
            // Deselect nodes
            TreeNode[] selected_nodes2 = selected_nodes.ToArray();
            foreach (TreeNode node in selected_nodes2)
                if (!nodes.Contains(node)) DeselectNode(node);
            // Select nodes
            foreach (TreeNode node in nodes)
                if (!selected_nodes.Contains(node)) SelectNode(node);
        }

        private void SelectBetweenNodes(TreeNode node1, TreeNode node2)
        {
            if (node1 == null || node1 == node2)
            {
                SelectSingleNode(node2);
            }
            else
            {
                List<TreeNode> nodes = new List<TreeNode>();
                TreeNode start;
                TreeNode stop;
                if (CompareNodes(node1, node2) < 0)
                {
                    start = node1;
                    stop = node2;
                }
                else
                {
                    start = node2;
                    stop = node1;
                }
                nodes.Add(start);
                while (start.NextVisibleNode != null && start.NextVisibleNode != stop)
                {
                    start = start.NextVisibleNode;
                    nodes.Add(start);
                }
                nodes.Add(stop);
                SelectNodes(nodes);
            }
        }

        private int CompareNodes(TreeNode node1, TreeNode node2)
        {
            TreeNode temp1 = node1;
            TreeNode temp2 = node2;
            while (temp1.Level > temp2.Level)
                temp1 = temp1.Parent;
            while (temp2.Level > temp1.Level)
                temp2 = temp2.Parent;
            while (temp1.Parent != temp2.Parent)
            {
                temp1 = temp1.Parent;
                temp2 = temp2.Parent;
            }
            if (temp1.Index == temp2.Index)
                if (node1 == node2)
                    return 0;
                else if (node1.Level < node2.Level)
                    return -1;
                else
                    return 1;
            else if (temp1.Index < temp2.Index)
                return -1;
            else
                return 1;
        }

        private void SortSelectedNodes()
        {
            selected_nodes.Sort(CompareNodes);
        }
    }
}
