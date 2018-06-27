using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.WordApi;

namespace NetOffice.WordApi.Behind
{
	/// <summary>
	/// DispatchInterface Window
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838990.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Window : COMObject, NetOffice.WordApi.Window
	{
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.WordApi.Window);
                return _contractType;
            }
        }
        private static Type _contractType;


		/// <summary>
		/// Instance Type
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
		public override Type InstanceType
		{
			get
			{
				return LateBindingApiWrapperType;
			}
		}

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(Window);                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Window() : base()
		{

		}

		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197003.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", typeof(NetOffice.WordApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839086.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Creator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838879.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822152.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Pane ActivePane
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Pane>(this, "ActivePane", typeof(NetOffice.WordApi.Pane));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835485.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Document Document
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Document>(this, "Document", typeof(NetOffice.WordApi.Document));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838919.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Panes Panes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Panes>(this, "Panes", typeof(NetOffice.WordApi.Panes));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845511.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Selection Selection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Selection>(this, "Selection", typeof(NetOffice.WordApi.Selection));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834260.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Left
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Left");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Left", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193092.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Top
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Top");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Top", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845639.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Width
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Width");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Width", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835119.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Height
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Height");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Height", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834813.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool Split
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Split");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Split", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839134.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 SplitVertical
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SplitVertical");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SplitVertical", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822965.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string Caption
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Caption");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Caption", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845378.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdWindowState WindowState
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdWindowState>(this, "WindowState");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "WindowState", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195421.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool DisplayRulers
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayRulers");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayRulers", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835761.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool DisplayVerticalRuler
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayVerticalRuler");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayVerticalRuler", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838505.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.View View
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.View>(this, "View", typeof(NetOffice.WordApi.View));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197875.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdWindowType Type
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdWindowType>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192589.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Window Next
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Window>(this, "Next", typeof(NetOffice.WordApi.Window));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196868.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Window Previous
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Window>(this, "Previous", typeof(NetOffice.WordApi.Window));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835402.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 WindowNumber
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "WindowNumber");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837323.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool DisplayVerticalScrollBar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayVerticalScrollBar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayVerticalScrollBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837924.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool DisplayHorizontalScrollBar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayHorizontalScrollBar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayHorizontalScrollBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192230.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single StyleAreaWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "StyleAreaWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StyleAreaWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840897.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool DisplayScreenTips
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayScreenTips");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayScreenTips", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191789.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 HorizontalPercentScrolled
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "HorizontalPercentScrolled");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HorizontalPercentScrolled", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844796.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 VerticalPercentScrolled
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "VerticalPercentScrolled");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "VerticalPercentScrolled", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839774.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool DocumentMap
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DocumentMap");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DocumentMap", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822144.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool Active
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Active");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 DocumentMapPercentWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "DocumentMapPercentWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DocumentMapPercentWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194852.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Index
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192140.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdIMEMode IMEMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdIMEMode>(this, "IMEMode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "IMEMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195060.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 UsableWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "UsableWidth");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821418.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 UsableHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "UsableHeight");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838517.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool EnvelopeVisible
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EnvelopeVisible");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnvelopeVisible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835783.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool DisplayRightRuler
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayRightRuler");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayRightRuler", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195605.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool DisplayLeftScrollBar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayLeftScrollBar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayLeftScrollBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820939.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool Visible
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Visible");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Visible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192623.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual bool Thumbnails
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Thumbnails");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Thumbnails", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197684.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdShowSourceDocuments ShowSourceDocuments
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdShowSourceDocuments>(this, "ShowSourceDocuments");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ShowSourceDocuments", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231484.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public virtual Int32 Hwnd
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Hwnd");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838523.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Activate()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Activate");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845707.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="routeDocument">optional object routeDocument</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Close(object saveChanges, object routeDocument)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close", saveChanges, routeDocument);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845707.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Close()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845707.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Close(object saveChanges)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close", saveChanges);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193698.aspx </remarks>
		/// <param name="down">optional object down</param>
		/// <param name="up">optional object up</param>
		/// <param name="toRight">optional object toRight</param>
		/// <param name="toLeft">optional object toLeft</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void LargeScroll(object down, object up, object toRight, object toLeft)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LargeScroll", down, up, toRight, toLeft);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193698.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void LargeScroll()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LargeScroll");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193698.aspx </remarks>
		/// <param name="down">optional object down</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void LargeScroll(object down)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LargeScroll", down);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193698.aspx </remarks>
		/// <param name="down">optional object down</param>
		/// <param name="up">optional object up</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void LargeScroll(object down, object up)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LargeScroll", down, up);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193698.aspx </remarks>
		/// <param name="down">optional object down</param>
		/// <param name="up">optional object up</param>
		/// <param name="toRight">optional object toRight</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void LargeScroll(object down, object up, object toRight)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LargeScroll", down, up, toRight);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193450.aspx </remarks>
		/// <param name="down">optional object down</param>
		/// <param name="up">optional object up</param>
		/// <param name="toRight">optional object toRight</param>
		/// <param name="toLeft">optional object toLeft</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SmallScroll(object down, object up, object toRight, object toLeft)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SmallScroll", down, up, toRight, toLeft);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193450.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SmallScroll()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SmallScroll");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193450.aspx </remarks>
		/// <param name="down">optional object down</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SmallScroll(object down)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SmallScroll", down);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193450.aspx </remarks>
		/// <param name="down">optional object down</param>
		/// <param name="up">optional object up</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SmallScroll(object down, object up)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SmallScroll", down, up);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193450.aspx </remarks>
		/// <param name="down">optional object down</param>
		/// <param name="up">optional object up</param>
		/// <param name="toRight">optional object toRight</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SmallScroll(object down, object up, object toRight)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SmallScroll", down, up, toRight);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840287.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Window NewWindow()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Window>(this, "NewWindow", typeof(NetOffice.WordApi.Window));
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOutOld()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOutOld(object background)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", background);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOutOld(object background, object append)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", background, append);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOutOld(object background, object append, object range)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", background, append, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOutOld(object background, object append, object range, object outputFileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", background, append, range, outputFileName);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839516.aspx </remarks>
		/// <param name="down">optional object down</param>
		/// <param name="up">optional object up</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PageScroll(object down, object up)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PageScroll", down, up);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839516.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PageScroll()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PageScroll");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839516.aspx </remarks>
		/// <param name="down">optional object down</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PageScroll(object down)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PageScroll", down);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838905.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SetFocus()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetFocus");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192575.aspx </remarks>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual object RangeFromPoint(Int32 x, Int32 y)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "RangeFromPoint", x, y);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836655.aspx </remarks>
		/// <param name="obj">object obj</param>
		/// <param name="start">optional object start</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void ScrollIntoView(object obj, object start)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ScrollIntoView", obj, start);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836655.aspx </remarks>
		/// <param name="obj">object obj</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void ScrollIntoView(object obj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ScrollIntoView", obj);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836626.aspx </remarks>
		/// <param name="screenPixelsLeft">Int32 screenPixelsLeft</param>
		/// <param name="screenPixelsTop">Int32 screenPixelsTop</param>
		/// <param name="screenPixelsWidth">Int32 screenPixelsWidth</param>
		/// <param name="screenPixelsHeight">Int32 screenPixelsHeight</param>
		/// <param name="obj">object obj</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void GetPoint(out Int32 screenPixelsLeft, out Int32 screenPixelsTop, out Int32 screenPixelsWidth, out Int32 screenPixelsHeight, object obj)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true,true,true,false);
			screenPixelsLeft = 0;
			screenPixelsTop = 0;
			screenPixelsWidth = 0;
			screenPixelsHeight = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(screenPixelsLeft, screenPixelsTop, screenPixelsWidth, screenPixelsHeight, obj);
			Invoker.Method(this, "GetPoint", paramsArray, modifiers);
			screenPixelsLeft = (Int32)paramsArray[0];
			screenPixelsTop = (Int32)paramsArray[1];
			screenPixelsWidth = (Int32)paramsArray[2];
			screenPixelsHeight = (Int32)paramsArray[3];
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
		/// <param name="printZoomPaperHeight">optional object printZoomPaperHeight</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth, object printZoomPaperHeight)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth, printZoomPaperHeight });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx </remarks>
		/// <param name="background">optional object background</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", background);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background, object append)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", background, append);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background, object append, object range)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", background, append, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background, object append, object range, object outputFileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", background, append, range, outputFileName);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
		/// <param name="printZoomPaperHeight">optional object printZoomPaperHeight</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth, object printZoomPaperHeight)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth, printZoomPaperHeight });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object background)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", background);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object background, object append)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", background, append);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object background, object append, object range)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", background, append, range);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", background, append, range, outputFileName);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void ToggleShowAllReviewers()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ToggleShowAllReviewers");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835142.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void ToggleRibbon()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ToggleRibbon");
		}

		#endregion

		#pragma warning restore
	}
}


