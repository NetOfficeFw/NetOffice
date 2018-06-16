using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSComctlLibApi
{
	#region Delegates

	#pragma warning disable
	public delegate void TabStrip_ClickEventHandler();
	public delegate void TabStrip_KeyDownEventHandler(ref Int16 keyCode, Int16 shift);
	public delegate void TabStrip_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void TabStrip_KeyUpEventHandler(ref Int16 keyCode, Int16 shift);
	public delegate void TabStrip_MouseDownEventHandler(Int16 button, Int16 shift, Int32 x, Int32 y);
	public delegate void TabStrip_MouseMoveEventHandler(Int16 button, Int16 shift, Int32 x, Int32 y);
	public delegate void TabStrip_MouseUpEventHandler(Int16 button, Int16 shift, Int32 x, Int32 y);
	public delegate void TabStrip_BeforeClickEventHandler(ref Int16 cancel);
	public delegate void TabStrip_OLEStartDragEventHandler(ref NetOffice.MSComctlLibApi.DataObject data, ref Int32 allowedEffects);
	public delegate void TabStrip_OLEGiveFeedbackEventHandler(ref Int32 effect, ref bool defaultCursors);
	public delegate void TabStrip_OLESetDataEventHandler(ref NetOffice.MSComctlLibApi.DataObject Data, ref Int16 dataFormat);
	public delegate void TabStrip_OLECompleteDragEventHandler(ref Int32 effect);
	public delegate void TabStrip_OLEDragOverEventHandler(ref NetOffice.MSComctlLibApi.DataObject data, ref Int32 effect, ref Int16 button, ref Int16 shift, ref Single x, ref Single y, ref Int16 state);
	public delegate void TabStrip_OLEDragDropEventHandler(ref NetOffice.MSComctlLibApi.DataObject data, ref Int32 effect, ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass TabStrip 
    /// SupportByVersion MSComctlLib, 6
    /// </summary>
    [SupportByVersion("MSComctlLib", 6)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ITabStripEvents))]
	[TypeId("1EFB6596-857C-11D1-B16A-00C0F0283628")]
    public interface TabStrip : ITabStrip, IEventBinding
    {
        #region Events
        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TabStrip_ClickEventHandler ClickEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TabStrip_KeyDownEventHandler KeyDownEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TabStrip_KeyPressEventHandler KeyPressEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TabStrip_KeyUpEventHandler KeyUpEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TabStrip_MouseDownEventHandler MouseDownEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TabStrip_MouseMoveEventHandler MouseMoveEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TabStrip_MouseUpEventHandler MouseUpEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TabStrip_BeforeClickEventHandler BeforeClickEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TabStrip_OLEStartDragEventHandler OLEStartDragEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TabStrip_OLEGiveFeedbackEventHandler OLEGiveFeedbackEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TabStrip_OLESetDataEventHandler OLESetDataEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TabStrip_OLECompleteDragEventHandler OLECompleteDragEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TabStrip_OLEDragOverEventHandler OLEDragOverEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event TabStrip_OLEDragDropEventHandler OLEDragDropEvent;

        #endregion
    }
}
