using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSComctlLibApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Toolbar_ButtonClickEventHandler(NetOffice.MSComctlLibApi.Button button);
	public delegate void Toolbar_ChangeEventHandler();
	public delegate void Toolbar_ClickEventHandler();
	public delegate void Toolbar_MouseDownEventHandler(Int16 button, Int16 shift, Int32 x, Int32 y);
	public delegate void Toolbar_MouseMoveEventHandler(Int16 button, Int16 shift, Int32 x, Int32 y);
	public delegate void Toolbar_MouseUpEventHandler(Int16 button, Int16 shift, Int32 x, Int32 y);
	public delegate void Toolbar_DblClickEventHandler();
	public delegate void Toolbar_OLEStartDragEventHandler(ref NetOffice.MSComctlLibApi.DataObject data, ref Int32 allowedEffects);
	public delegate void Toolbar_OLEGiveFeedbackEventHandler(ref Int32 effect, ref bool defaultCursors);
	public delegate void Toolbar_OLESetDataEventHandler(ref NetOffice.MSComctlLibApi.DataObject data, ref Int16 dataFormat);
	public delegate void Toolbar_OLECompleteDragEventHandler(ref Int32 effect);
	public delegate void Toolbar_OLEDragOverEventHandler(ref NetOffice.MSComctlLibApi.DataObject data, ref Int32 effect, ref Int16 button, ref Int16 shift, ref Single x, ref Single y, ref Int16 state);
	public delegate void Toolbar_OLEDragDropEventHandler(ref NetOffice.MSComctlLibApi.DataObject data, ref Int32 effect, ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void Toolbar_ButtonMenuClickEventHandler(NetOffice.MSComctlLibApi.ButtonMenu buttonMenu);
	public delegate void Toolbar_ButtonDropDownEventHandler(NetOffice.MSComctlLibApi.Button button);
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass Toolbar 
    /// SupportByVersion MSComctlLib, 6
    /// </summary>
    [SupportByVersion("MSComctlLib", 6)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.IToolbarEvents))]
	[TypeId("66833FE6-8583-11D1-B16A-00C0F0283628")]
    public interface Toolbar : IToolbar, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event Toolbar_ButtonClickEventHandler ButtonClickEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event Toolbar_ChangeEventHandler ChangeEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event Toolbar_ClickEventHandler ClickEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event Toolbar_MouseDownEventHandler MouseDownEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event Toolbar_MouseMoveEventHandler MouseMoveEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event Toolbar_MouseUpEventHandler MouseUpEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event Toolbar_DblClickEventHandler DblClickEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event Toolbar_OLEStartDragEventHandler OLEStartDragEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event Toolbar_OLEGiveFeedbackEventHandler OLEGiveFeedbackEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event Toolbar_OLESetDataEventHandler OLESetDataEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event Toolbar_OLECompleteDragEventHandler OLECompleteDragEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event Toolbar_OLEDragOverEventHandler OLEDragOverEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event Toolbar_OLEDragDropEventHandler OLEDragDropEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event Toolbar_ButtonMenuClickEventHandler ButtonMenuClickEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        event Toolbar_ButtonDropDownEventHandler ButtonDropDownEvent;

        #endregion
    }
}
