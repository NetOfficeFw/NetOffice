using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.DeveloperToolbox.Controls.Tree
{
	public class TreeGridNodeEventBase
	{
		private TreeGridNode _node;

		public TreeGridNodeEventBase(TreeGridNode node)
		{
			this._node = node;
		}

		public TreeGridNode Node
		{
			get { return _node; }
		}
	}
	public class CollapsingEventArgs : System.ComponentModel.CancelEventArgs
	{
		private TreeGridNode _node;

		private CollapsingEventArgs() { }
		public CollapsingEventArgs(TreeGridNode node)
			: base()
		{
			this._node = node;
		}
		public TreeGridNode Node
		{
			get { return _node; }
		}

	}
	public class CollapsedEventArgs : TreeGridNodeEventBase
	{
		public CollapsedEventArgs(TreeGridNode node)
			: base(node)
		{
		}
	}

	public class ExpandingEventArgs:System.ComponentModel.CancelEventArgs
	{
		private TreeGridNode _node;

		private ExpandingEventArgs() { }
		public ExpandingEventArgs(TreeGridNode node):base()
		{
			this._node = node;
		}
		public TreeGridNode Node
		{
			get { return _node; }
		}

	}
	public class ExpandedEventArgs : TreeGridNodeEventBase
	{
		public ExpandedEventArgs(TreeGridNode node):base(node)
		{
		}
	}

	public delegate void ExpandingEventHandler(object sender, ExpandingEventArgs e);
	public delegate void ExpandedEventHandler(object sender, ExpandedEventArgs e);

	public delegate void CollapsingEventHandler(object sender, CollapsingEventArgs e);
	public delegate void CollapsedEventHandler(object sender, CollapsedEventArgs e);

}
