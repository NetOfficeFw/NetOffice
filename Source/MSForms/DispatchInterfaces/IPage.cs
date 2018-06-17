using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	/// <summary>
	/// DispatchInterface IPage 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("5CEF5613-713D-11CE-80C9-00AA00611080")]
    [CoClassSource(typeof(NetOffice.MSFormsApi.Page))]
    public interface IPage : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Controls Controls { get; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.MSFormsApi.Controls Selected { get; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.MSFormsApi.Control ActiveControl { get; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool CanPaste { get; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool CanRedo { get; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool CanUndo { get; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Enums.fmCycle Cycle { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		string Caption { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		string ControlTipText { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		bool Enabled { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Int32 Index { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Single InsideHeight { get; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Single InsideWidth { get; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Enums.fmScrollBars KeepScrollBarsVisible { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Enums.fmPictureAlignment PictureAlignment { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2), NativeResult]
		stdole.Picture Picture { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Enums.fmPictureSizeMode PictureSizeMode { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		bool PictureTiling { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Enums.fmScrollBars ScrollBars { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Single ScrollHeight { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Single ScrollLeft { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Single ScrollTop { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Single ScrollWidth { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		string Tag { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Enums.fmTransitionEffect TransitionEffect { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Int32 TransitionPeriod { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.MSFormsApi.Enums.fmVerticalScrollBarSide VerticalScrollBarSide { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		bool Visible { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Int16 Zoom { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.MSFormsApi.Enums.fmMode DesignMode { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.MSFormsApi.Enums.fmMode ShowToolbox { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.MSFormsApi.Enums.fmMode ShowGridDots { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.MSFormsApi.Enums.fmMode SnapToGrid { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Single GridX { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Single GridY { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		string Accelerator { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSForms", 2), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Parent { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="insideHeight">Int32 insideHeight</param>
		[SupportByVersion("MSForms", 2)]
		void _GetInsideHeight(out Int32 insideHeight);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="insideWidth">Int32 insideWidth</param>
		[SupportByVersion("MSForms", 2)]
		void _GetInsideWidth(out Int32 insideWidth);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="scrollHeight">Int32 scrollHeight</param>
		[SupportByVersion("MSForms", 2)]
		void _SetScrollHeight(Int32 scrollHeight);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="scrollHeight">Int32 scrollHeight</param>
		[SupportByVersion("MSForms", 2)]
		void _GetScrollHeight(out Int32 scrollHeight);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="scrollLeft">Int32 scrollLeft</param>
		[SupportByVersion("MSForms", 2)]
		void _SetScrollLeft(Int32 scrollLeft);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="scrollLeft">Int32 scrollLeft</param>
		[SupportByVersion("MSForms", 2)]
		void _GetScrollLeft(out Int32 scrollLeft);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="scrollTop">Int32 scrollTop</param>
		[SupportByVersion("MSForms", 2)]
		void _SetScrollTop(Int32 scrollTop);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="scrollTop">Int32 scrollTop</param>
		[SupportByVersion("MSForms", 2)]
		void _GetScrollTop(out Int32 scrollTop);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="scrollWidth">Int32 scrollWidth</param>
		[SupportByVersion("MSForms", 2)]
		void _SetScrollWidth(Int32 scrollWidth);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="scrollWidth">Int32 scrollWidth</param>
		[SupportByVersion("MSForms", 2)]
		void _GetScrollWidth(out Int32 scrollWidth);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="gridX">Int32 gridX</param>
		[SupportByVersion("MSForms", 2)]
		void _SetGridX(Int32 gridX);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="gridX">Int32 gridX</param>
		[SupportByVersion("MSForms", 2)]
		void _GetGridX(out Int32 gridX);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="gridY">Int32 gridY</param>
		[SupportByVersion("MSForms", 2)]
		void _SetGridY(Int32 gridY);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="gridY">Int32 gridY</param>
		[SupportByVersion("MSForms", 2)]
		void _GetGridY(out Int32 gridY);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		void Copy();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		void Cut();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		void Paste();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		void RedoAction();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		void Repaint();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="xAction">optional object xAction</param>
		/// <param name="yAction">optional object yAction</param>
		[SupportByVersion("MSForms", 2)]
		void Scroll(object xAction, object yAction);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		void Scroll();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="xAction">optional object xAction</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		void Scroll(object xAction);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		void SetDefaultTabOrder();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		void UndoAction();

		#endregion
	}
}
