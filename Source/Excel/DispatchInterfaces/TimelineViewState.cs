using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface TimelineViewState 
	/// SupportByVersion Excel, 15, 16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.timelineviewstate"/> </remarks>
	[SupportByVersion("Excel", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class TimelineViewState : COMObject
	{
		#pragma warning disable

		#region Type Information

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
                    _type = typeof(TimelineViewState);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public TimelineViewState(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public TimelineViewState(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TimelineViewState(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TimelineViewState(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TimelineViewState(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TimelineViewState(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TimelineViewState() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TimelineViewState(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.timelineviewstate.application"/> </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.timelineviewstate.creator"/> </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.timelineviewstate.parent"/> </remarks>
		[SupportByVersion("Excel", 15, 16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.timelineviewstate.showheader"/> </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public bool ShowHeader
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowHeader");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowHeader", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.timelineviewstate.showselectionlabel"/> </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public bool ShowSelectionLabel
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowSelectionLabel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowSelectionLabel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.timelineviewstate.showtimelevel"/> </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public bool ShowTimeLevel
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowTimeLevel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowTimeLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.timelineviewstate.showhorizontalscrollbar"/> </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public bool ShowHorizontalScrollbar
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowHorizontalScrollbar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowHorizontalScrollbar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.timelineviewstate.level"/> </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Enums.XlTimelineLevel Level
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlTimelineLevel>(this, "Level");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Level", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
