using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface Slides 
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746073.aspx </remarks>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), HasIndexProperty(IndexInvoke.Method, "Item")]
 	public class Slides : Collection
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
                    _type = typeof(Slides);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Slides(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Slides(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Slides(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Slides(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Slides(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Slides(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Slides() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Slides(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745043.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745238.aspx </remarks>
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
		public NetOffice.PowerPointApi.Slide this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Slide>(this, "Item", NetOffice.PowerPointApi.Slide.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744787.aspx </remarks>
		/// <param name="slideID">Int32 slideID</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Slide FindBySlideID(Int32 slideID)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Slide>(this, "FindBySlideID", NetOffice.PowerPointApi.Slide.LateBindingApiWrapperType, slideID);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		/// <param name="layout">NetOffice.PowerPointApi.Enums.PpSlideLayout layout</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Slide Add(Int32 index, NetOffice.PowerPointApi.Enums.PpSlideLayout layout)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Slide>(this, "Add", NetOffice.PowerPointApi.Slide.LateBindingApiWrapperType, index, layout);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746047.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="index">Int32 index</param>
		/// <param name="slideStart">optional Int32 SlideStart = 1</param>
		/// <param name="slideEnd">optional Int32 SlideEnd = -1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Int32 InsertFromFile(string fileName, Int32 index, object slideStart, object slideEnd)
		{
			return Factory.ExecuteInt32MethodGet(this, "InsertFromFile", fileName, index, slideStart, slideEnd);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746047.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="index">Int32 index</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Int32 InsertFromFile(string fileName, Int32 index)
		{
			return Factory.ExecuteInt32MethodGet(this, "InsertFromFile", fileName, index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746047.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="index">Int32 index</param>
		/// <param name="slideStart">optional Int32 SlideStart = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Int32 InsertFromFile(string fileName, Int32 index, object slideStart)
		{
			return Factory.ExecuteInt32MethodGet(this, "InsertFromFile", fileName, index, slideStart);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746710.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.SlideRange Range(object index)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.SlideRange>(this, "Range", NetOffice.PowerPointApi.SlideRange.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746710.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.SlideRange Range()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.SlideRange>(this, "Range", NetOffice.PowerPointApi.SlideRange.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744391.aspx </remarks>
		/// <param name="index">optional Int32 index = -1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.SlideRange Paste(object index)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.SlideRange>(this, "Paste", NetOffice.PowerPointApi.SlideRange.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744391.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.SlideRange Paste()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.SlideRange>(this, "Paste", NetOffice.PowerPointApi.SlideRange.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746586.aspx </remarks>
		/// <param name="index">Int32 index</param>
		/// <param name="pCustomLayout">NetOffice.PowerPointApi.CustomLayout pCustomLayout</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.Slide AddSlide(Int32 index, NetOffice.PowerPointApi.CustomLayout pCustomLayout)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Slide>(this, "AddSlide", NetOffice.PowerPointApi.Slide.LateBindingApiWrapperType, index, pCustomLayout);
		}

		#endregion

		#pragma warning restore
	}
}
