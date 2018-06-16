using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSComctlLibApi
{
	#region Delegates

	#pragma warning disable
	public delegate void ProgressBar_MouseDownEventHandler(Int16 button, Int16 shift, Int32 x, Int32 y);
	public delegate void ProgressBar_MouseMoveEventHandler(Int16 button, Int16 shift, Int32 x, Int32 y);
	public delegate void ProgressBar_MouseUpEventHandler(Int16 button, Int16 shift, Int32 x, Int32 y);
	public delegate void ProgressBar_ClickEventHandler();
	public delegate void ProgressBar_OLEStartDragEventHandler(ref NetOffice.MSComctlLibApi.DataObject data, ref Int32 allowedEffects);
	public delegate void ProgressBar_OLEGiveFeedbackEventHandler(ref Int32 effect, ref bool defaultCursors);
	public delegate void ProgressBar_OLESetDataEventHandler(ref NetOffice.MSComctlLibApi.DataObject data, ref Int16 dataFormat);
	public delegate void ProgressBar_OLECompleteDragEventHandler(ref Int32 effect);
	public delegate void ProgressBar_OLEDragOverEventHandler(ref NetOffice.MSComctlLibApi.DataObject data, ref Int32 effect, ref Int16 button, ref Int16 shift, ref Single x, ref Single y, ref Int16 state);
	public delegate void ProgressBar_OLEDragDropEventHandler(ref NetOffice.MSComctlLibApi.DataObject data, ref Int32 effect, ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass ProgressBar 
    /// SupportByVersion MSComctlLib, 6
    /// </summary>
    [SupportByVersion("MSComctlLib", 6)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.IProgressBarEvents))]
	[TypeId("35053A22-8589-11D1-B16A-00C0F0283628")]
    public interface ProgressBar : IProgressBar, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ProgressBar_MouseDownEventHandler MouseDownEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ProgressBar_MouseMoveEventHandler MouseMoveEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ProgressBar_MouseUpEventHandler MouseUpEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ProgressBar_ClickEventHandler ClickEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ProgressBar_OLEStartDragEventHandler OLEStartDragEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ProgressBar_OLEGiveFeedbackEventHandler OLEGiveFeedbackEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ProgressBar_OLESetDataEventHandler OLESetDataEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ProgressBar_OLECompleteDragEventHandler OLECompleteDragEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ProgressBar_OLEDragOverEventHandler OLEDragOverEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event ProgressBar_OLEDragDropEventHandler OLEDragDropEvent;

        #endregion
    }
}
