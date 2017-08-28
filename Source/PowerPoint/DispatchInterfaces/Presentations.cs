using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface Presentations 
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743968.aspx </remarks>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), HasIndexProperty(IndexInvoke.Method, "Item")]
 	public class Presentations : Collection
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
                    _type = typeof(Presentations);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Presentations(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Presentations(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Presentations(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Presentations(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Presentations(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Presentations(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Presentations() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Presentations(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744974.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Application>(this, "Application", NetOffice.PowerPointApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744972.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.PowerPointApi.Presentation this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Item", NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745733.aspx </remarks>
		/// <param name="withWindow">optional NetOffice.OfficeApi.Enums.MsoTriState WithWindow = -1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Presentation Add(object withWindow)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Add", NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType, withWindow);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745733.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Presentation Add()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Add", NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746171.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="readOnly">optional NetOffice.OfficeApi.Enums.MsoTriState ReadOnly = 0</param>
		/// <param name="untitled">optional NetOffice.OfficeApi.Enums.MsoTriState Untitled = 0</param>
		/// <param name="withWindow">optional NetOffice.OfficeApi.Enums.MsoTriState WithWindow = -1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Presentation Open(string fileName, object readOnly, object untitled, object withWindow)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Open", NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType, fileName, readOnly, untitled, withWindow);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746171.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Presentation Open(string fileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Open", NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType, fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746171.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="readOnly">optional NetOffice.OfficeApi.Enums.MsoTriState ReadOnly = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Presentation Open(string fileName, object readOnly)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Open", NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType, fileName, readOnly);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746171.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="readOnly">optional NetOffice.OfficeApi.Enums.MsoTriState ReadOnly = 0</param>
		/// <param name="untitled">optional NetOffice.OfficeApi.Enums.MsoTriState Untitled = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Presentation Open(string fileName, object readOnly, object untitled)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Open", NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType, fileName, readOnly, untitled);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="readOnly">optional NetOffice.OfficeApi.Enums.MsoTriState ReadOnly = 0</param>
		/// <param name="untitled">optional NetOffice.OfficeApi.Enums.MsoTriState Untitled = 0</param>
		/// <param name="withWindow">optional NetOffice.OfficeApi.Enums.MsoTriState WithWindow = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Presentation OpenOld(string fileName, object readOnly, object untitled, object withWindow)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "OpenOld", NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType, fileName, readOnly, untitled, withWindow);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Presentation OpenOld(string fileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "OpenOld", NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType, fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="readOnly">optional NetOffice.OfficeApi.Enums.MsoTriState ReadOnly = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Presentation OpenOld(string fileName, object readOnly)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "OpenOld", NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType, fileName, readOnly);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="readOnly">optional NetOffice.OfficeApi.Enums.MsoTriState ReadOnly = 0</param>
		/// <param name="untitled">optional NetOffice.OfficeApi.Enums.MsoTriState Untitled = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Presentation OpenOld(string fileName, object readOnly, object untitled)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "OpenOld", NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType, fileName, readOnly, untitled);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746209.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void CheckOut(string fileName)
		{
			 Factory.ExecuteMethod(this, "CheckOut", fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745034.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public bool CanCheckOut(string fileName)
		{
			return Factory.ExecuteBoolMethodGet(this, "CanCheckOut", fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744741.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="readOnly">optional NetOffice.OfficeApi.Enums.MsoTriState ReadOnly = 0</param>
		/// <param name="untitled">optional NetOffice.OfficeApi.Enums.MsoTriState Untitled = 0</param>
		/// <param name="withWindow">optional NetOffice.OfficeApi.Enums.MsoTriState WithWindow = -1</param>
		/// <param name="openAndRepair">optional NetOffice.OfficeApi.Enums.MsoTriState OpenAndRepair = 0</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.Presentation Open2007(string fileName, object readOnly, object untitled, object withWindow, object openAndRepair)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Open2007", NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType, new object[]{ fileName, readOnly, untitled, withWindow, openAndRepair });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744741.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.Presentation Open2007(string fileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Open2007", NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType, fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744741.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="readOnly">optional NetOffice.OfficeApi.Enums.MsoTriState ReadOnly = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.Presentation Open2007(string fileName, object readOnly)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Open2007", NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType, fileName, readOnly);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744741.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="readOnly">optional NetOffice.OfficeApi.Enums.MsoTriState ReadOnly = 0</param>
		/// <param name="untitled">optional NetOffice.OfficeApi.Enums.MsoTriState Untitled = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.Presentation Open2007(string fileName, object readOnly, object untitled)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Open2007", NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType, fileName, readOnly, untitled);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744741.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="readOnly">optional NetOffice.OfficeApi.Enums.MsoTriState ReadOnly = 0</param>
		/// <param name="untitled">optional NetOffice.OfficeApi.Enums.MsoTriState Untitled = 0</param>
		/// <param name="withWindow">optional NetOffice.OfficeApi.Enums.MsoTriState WithWindow = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.Presentation Open2007(string fileName, object readOnly, object untitled, object withWindow)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Open2007", NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType, fileName, readOnly, untitled, withWindow);
		}

		#endregion

		#pragma warning restore
	}
}
