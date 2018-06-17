using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	/// <summary>
	/// DispatchInterface IControl 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("04598FC6-866C-11CF-AB7C-00AA00C08FCF")]
    [CoClassSource(typeof(NetOffice.MSFormsApi.Control))]
    public interface IControl : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		bool Cancel { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		string ControlSource { get; set; }

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
		bool Default { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Single Height { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Int32 HelpContextID { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool InSelection { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.MSFormsApi.Enums.fmLayoutEffect LayoutEffect { get; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Single Left { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Single OldHeight { get; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Single OldLeft { get; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Single OldTop { get; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Single OldWidth { get; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSForms", 2), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Object { get; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSForms", 2), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		string RowSource { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Int16 RowSourceType { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Int16 TabIndex { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		bool TabStop { get; set; }

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
		Single Top { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object BoundValue { get; set; }

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
		Single Width { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="height">Int32 height</param>
		[SupportByVersion("MSForms", 2)]
		void _SetHeight(Int32 height);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="height">Int32 height</param>
		[SupportByVersion("MSForms", 2)]
		void _GetHeight(out Int32 height);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="left">Int32 left</param>
		[SupportByVersion("MSForms", 2)]
		void _SetLeft(Int32 left);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="left">Int32 left</param>
		[SupportByVersion("MSForms", 2)]
		void _GetLeft(out Int32 left);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="oldHeight">Int32 oldHeight</param>
		[SupportByVersion("MSForms", 2)]
		void _GetOldHeight(out Int32 oldHeight);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="oldLeft">Int32 oldLeft</param>
		[SupportByVersion("MSForms", 2)]
		void _GetOldLeft(out Int32 oldLeft);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="oldTop">Int32 oldTop</param>
		[SupportByVersion("MSForms", 2)]
		void _GetOldTop(out Int32 oldTop);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="oldWidth">Int32 oldWidth</param>
		[SupportByVersion("MSForms", 2)]
		void _GetOldWidth(out Int32 oldWidth);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="top">Int32 top</param>
		[SupportByVersion("MSForms", 2)]
		void _SetTop(Int32 top);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="top">Int32 top</param>
		[SupportByVersion("MSForms", 2)]
		void _GetTop(out Int32 top);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="width">Int32 width</param>
		[SupportByVersion("MSForms", 2)]
		void _SetWidth(Int32 width);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="width">Int32 width</param>
		[SupportByVersion("MSForms", 2)]
		void _GetWidth(out Int32 width);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		/// <param name="layout">optional object layout</param>
		[SupportByVersion("MSForms", 2)]
		void Move(object left, object top, object width, object height, object layout);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		void Move();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="left">optional object left</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		void Move(object left);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		void Move(object left, object top);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		void Move(object left, object top, object width);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		void Move(object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="zPosition">optional object zPosition</param>
		[SupportByVersion("MSForms", 2)]
		void ZOrder(object zPosition);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		void ZOrder();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="selectInGroup">bool selectInGroup</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSForms", 2)]
		void Select(bool selectInGroup);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		void SetFocus();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Int32 _GethWnd();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Int32 _GetID();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="width">Int32 width</param>
		/// <param name="height">Int32 height</param>
		[SupportByVersion("MSForms", 2)]
		void _Move(Int32 left, Int32 top, Int32 width, Int32 height);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="zPosition">NetOffice.MSFormsApi.Enums.fmZOrder zPosition</param>
		[SupportByVersion("MSForms", 2)]
		void _ZOrder(NetOffice.MSFormsApi.Enums.fmZOrder zPosition);

		#endregion
	}
}
