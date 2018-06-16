using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSComctlLibApi
{
	#region Delegates

	#pragma warning disable
	public delegate void TreeView_BeforeLabelEditEventHandler(ref Int16 cancel);
	public delegate void TreeView_AfterLabelEditEventHandler(ref Int16 cancel, ref string newString);
	public delegate void TreeView_CollapseEventHandler(NetOffice.MSComctlLibApi.Node node);
	public delegate void TreeView_ExpandEventHandler(NetOffice.MSComctlLibApi.Node node);
	public delegate void TreeView_NodeClickEventHandler(NetOffice.MSComctlLibApi.Node node);
	public delegate void TreeView_KeyDownEventHandler(ref Int16 keyCode, Int16 shift);
	public delegate void TreeView_KeyUpEventHandler(ref Int16 keyCode, Int16 shift);
	public delegate void TreeView_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void TreeView_MouseDownEventHandler(Int16 button, Int16 shift, Int32 x, Int32 y);
	public delegate void TreeView_MouseMoveEventHandler(Int16 button, Int16 shift, Int32 x, Int32 y);
	public delegate void TreeView_MouseUpEventHandler(Int16 button, Int16 shift, Int32 x, Int32 y);
	public delegate void TreeView_ClickEventHandler();
	public delegate void TreeView_DblClickEventHandler();
	public delegate void TreeView_NodeCheckEventHandler(NetOffice.MSComctlLibApi.Node node);
	public delegate void TreeView_OLEStartDragEventHandler(ref NetOffice.MSComctlLibApi.DataObject data, ref Int32 allowedEffects);
	public delegate void TreeView_OLEGiveFeedbackEventHandler(ref Int32 effect, ref bool defaultCursors);
	public delegate void TreeView_OLESetDataEventHandler(ref NetOffice.MSComctlLibApi.DataObject data, ref Int16 dataFormat);
	public delegate void TreeView_OLECompleteDragEventHandler(ref Int32 effect);
	public delegate void TreeView_OLEDragOverEventHandler(ref NetOffice.MSComctlLibApi.DataObject data, ref Int32 effect, ref Int16 button, ref Int16 shift, ref Single x, ref Single y, ref Int16 state);
	public delegate void TreeView_OLEDragDropEventHandler(ref NetOffice.MSComctlLibApi.DataObject data, ref Int32 effect, ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass TreeView 
    /// SupportByVersion MSComctlLib, 6
    /// </summary>
    [SupportByVersion("MSComctlLib", 6)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ITreeViewEvents))]
	[TypeId("C74190B6-8589-11D1-B16A-00C0F0283628")]
    public interface TreeView : ITreeView, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TreeView_BeforeLabelEditEventHandler BeforeLabelEditEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TreeView_AfterLabelEditEventHandler AfterLabelEditEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TreeView_CollapseEventHandler CollapseEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TreeView_ExpandEventHandler ExpandEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TreeView_NodeClickEventHandler NodeClickEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TreeView_KeyDownEventHandler KeyDownEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TreeView_KeyUpEventHandler KeyUpEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TreeView_KeyPressEventHandler KeyPressEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TreeView_MouseDownEventHandler MouseDownEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TreeView_MouseMoveEventHandler MouseMoveEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TreeView_MouseUpEventHandler MouseUpEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TreeView_ClickEventHandler ClickEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TreeView_DblClickEventHandler DblClickEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TreeView_NodeCheckEventHandler NodeCheckEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TreeView_OLEStartDragEventHandler OLEStartDragEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TreeView_OLEGiveFeedbackEventHandler OLEGiveFeedbackEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TreeView_OLESetDataEventHandler OLESetDataEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TreeView_OLECompleteDragEventHandler OLECompleteDragEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TreeView_OLEDragOverEventHandler OLEDragOverEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TreeView_OLEDragDropEventHandler OLEDragDropEvent;

        #endregion
    }
}
