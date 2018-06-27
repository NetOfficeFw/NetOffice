using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.PowerPointApi;

namespace NetOffice.PowerPointApi.Behind
{
	/// <summary>
	/// DispatchInterface Presentations
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743968.aspx </remarks>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), HasIndexProperty(IndexInvoke.Method, "Item")]
 	public class Presentations : Collection, NetOffice.PowerPointApi.Presentations
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
                    _contractType = typeof(NetOffice.PowerPointApi.Presentations);
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
                    _type = typeof(Presentations);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Presentations() : base()
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
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Application>(this, "Application", typeof(NetOffice.PowerPointApi.Application));
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
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
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
				return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Item", typeof(NetOffice.PowerPointApi.Presentation), index);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Add", typeof(NetOffice.PowerPointApi.Presentation), withWindow);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745733.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Presentation Add()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Add", typeof(NetOffice.PowerPointApi.Presentation));
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Open", typeof(NetOffice.PowerPointApi.Presentation), fileName, readOnly, untitled, withWindow);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Open", typeof(NetOffice.PowerPointApi.Presentation), fileName);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Open", typeof(NetOffice.PowerPointApi.Presentation), fileName, readOnly);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Open", typeof(NetOffice.PowerPointApi.Presentation), fileName, readOnly, untitled);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "OpenOld", typeof(NetOffice.PowerPointApi.Presentation), fileName, readOnly, untitled, withWindow);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "OpenOld", typeof(NetOffice.PowerPointApi.Presentation), fileName);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "OpenOld", typeof(NetOffice.PowerPointApi.Presentation), fileName, readOnly);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "OpenOld", typeof(NetOffice.PowerPointApi.Presentation), fileName, readOnly, untitled);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746209.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void CheckOut(string fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckOut", fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745034.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public bool CanCheckOut(string fileName)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CanCheckOut", fileName);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Open2007", typeof(NetOffice.PowerPointApi.Presentation), new object[]{ fileName, readOnly, untitled, withWindow, openAndRepair });
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Open2007", typeof(NetOffice.PowerPointApi.Presentation), fileName);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Open2007", typeof(NetOffice.PowerPointApi.Presentation), fileName, readOnly);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Open2007", typeof(NetOffice.PowerPointApi.Presentation), fileName, readOnly, untitled);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Presentation>(this, "Open2007", typeof(NetOffice.PowerPointApi.Presentation), fileName, readOnly, untitled, withWindow);
		}

		#endregion

		#pragma warning restore
	}
}


