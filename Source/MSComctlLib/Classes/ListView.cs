using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSComctlLibApi
{
	#region Delegates

	#pragma warning disable
	public delegate void ListView_BeforeLabelEditEventHandler(ref Int16 cancel);
	public delegate void ListView_AfterLabelEditEventHandler(ref Int16 cancel, ref string newString);
	public delegate void ListView_ColumnClickEventHandler(NetOffice.MSComctlLibApi.ColumnHeader columnHeader);
	public delegate void ListView_ItemClickEventHandler(NetOffice.MSComctlLibApi.ListItem item);
	public delegate void ListView_KeyDownEventHandler(ref Int16 keyCode, Int16 shift);
	public delegate void ListView_KeyUpEventHandler(ref Int16 keyCode, Int16 shift);
	public delegate void ListView_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void ListView_MouseDownEventHandler(Int16 button, Int16 shift, Int32 x, Int32 y);
	public delegate void ListView_MouseMoveEventHandler(Int16 button, Int16 shift, Int32 x, Int32 y);
	public delegate void ListView_MouseUpEventHandler(Int16 button, Int16 shift, Int32 x, Int32 y);
	public delegate void ListView_ClickEventHandler();
	public delegate void ListView_DblClickEventHandler();
	public delegate void ListView_OLEStartDragEventHandler(ref NetOffice.MSComctlLibApi.DataObject Data, ref Int32 allowedEffects);
	public delegate void ListView_OLEGiveFeedbackEventHandler(ref Int32 effect, ref bool defaultCursors);
	public delegate void ListView_OLESetDataEventHandler(ref NetOffice.MSComctlLibApi.DataObject data, ref Int16 dataFormat);
	public delegate void ListView_OLECompleteDragEventHandler(ref Int32 effect);
	public delegate void ListView_OLEDragOverEventHandler(ref NetOffice.MSComctlLibApi.DataObject data, ref Int32 effect, ref Int16 button, ref Int16 shift, ref Single x, ref Single y, ref Int16 state);
	public delegate void ListView_OLEDragDropEventHandler(ref NetOffice.MSComctlLibApi.DataObject data, ref Int32 effect, ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void ListView_ItemCheckEventHandler(NetOffice.MSComctlLibApi.ListItem item);
    #pragma warning restore

    #endregion
    
    /// <summary>
    /// CoClass ListView 
    /// SupportByVersion MSComctlLib, 6
    /// </summary>
    [SupportByVersion("MSComctlLib", 6)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ListViewEvents))]
	[TypeId("BDD1F04B-858B-11D1-B16A-00C0F0283628")]
    public interface ListView : IListView, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ListView_BeforeLabelEditEventHandler BeforeLabelEditEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ListView_AfterLabelEditEventHandler AfterLabelEditEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ListView_ColumnClickEventHandler ColumnClickEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ListView_ItemClickEventHandler ItemClickEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ListView_KeyDownEventHandler KeyDownEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ListView_KeyUpEventHandler KeyUpEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ListView_KeyPressEventHandler KeyPressEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ListView_MouseDownEventHandler MouseDownEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ListView_MouseMoveEventHandler MouseMoveEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ListView_MouseUpEventHandler MouseUpEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ListView_ClickEventHandler ClickEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ListView_DblClickEventHandler DblClickEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ListView_OLEStartDragEventHandler OLEStartDragEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ListView_OLEGiveFeedbackEventHandler OLEGiveFeedbackEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ListView_OLESetDataEventHandler OLESetDataEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ListView_OLECompleteDragEventHandler OLECompleteDragEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ListView_OLEDragOverEventHandler OLEDragOverEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ListView_OLEDragDropEventHandler OLEDragDropEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ListView_ItemCheckEventHandler ItemCheckEvent;

        #endregion
    }
}
