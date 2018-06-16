using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSComctlLibApi
{
	#region Delegates

	#pragma warning disable
	public delegate void StatusBar_PanelClickEventHandler(NetOffice.MSComctlLibApi.Panel panel);
	public delegate void StatusBar_PanelDblClickEventHandler(NetOffice.MSComctlLibApi.Panel panel);
	public delegate void StatusBar_MouseDownEventHandler(Int16 button, Int16 shift, Int32 x, Int32 y);
	public delegate void StatusBar_MouseMoveEventHandler(Int16 button, Int16 shift, Int32 x, Int32 y);
	public delegate void StatusBar_MouseUpEventHandler(Int16 button, Int16 shift, Int32 x, Int32 y);
	public delegate void StatusBar_ClickEventHandler();
	public delegate void StatusBar_DblClickEventHandler();
	public delegate void StatusBar_OLEStartDragEventHandler(ref NetOffice.MSComctlLibApi.DataObject data, ref Int32 allowedEffects);
	public delegate void StatusBar_OLEGiveFeedbackEventHandler(ref Int32 effect, ref bool defaultCursors);
	public delegate void StatusBar_OLESetDataEventHandler(ref NetOffice.MSComctlLibApi.DataObject data, ref Int16 dataFormat);
	public delegate void StatusBar_OLECompleteDragEventHandler(ref Int32 effect);
	public delegate void StatusBar_OLEDragOverEventHandler(ref NetOffice.MSComctlLibApi.DataObject data, ref Int32 effect, ref Int16 button, ref Int16 shift, ref Single x, ref Single y, ref Int16 state);
	public delegate void StatusBar_OLEDragDropEventHandler(ref NetOffice.MSComctlLibApi.DataObject data, ref Int32 effect, ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass StatusBar 
    /// SupportByVersion MSComctlLib, 6
    /// </summary>
    [SupportByVersion("MSComctlLib", 6)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.IStatusBarEvents))]
	[TypeId("8E3867A3-8586-11D1-B16A-00C0F0283628")]
    public interface StatusBar : IStatusBar, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event StatusBar_PanelClickEventHandler PanelClickEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event StatusBar_PanelDblClickEventHandler PanelDblClickEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event StatusBar_MouseDownEventHandler MouseDownEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event StatusBar_MouseMoveEventHandler MouseMoveEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event StatusBar_MouseUpEventHandler MouseUpEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event StatusBar_ClickEventHandler ClickEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event StatusBar_DblClickEventHandler DblClickEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event StatusBar_OLEStartDragEventHandler OLEStartDragEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event StatusBar_OLEGiveFeedbackEventHandler OLEGiveFeedbackEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event StatusBar_OLESetDataEventHandler OLESetDataEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event StatusBar_OLECompleteDragEventHandler OLECompleteDragEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event StatusBar_OLEDragOverEventHandler OLEDragOverEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event StatusBar_OLEDragDropEventHandler OLEDragDropEvent;

        #endregion
    }
}
