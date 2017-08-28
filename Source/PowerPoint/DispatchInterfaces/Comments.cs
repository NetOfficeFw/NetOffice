using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface Comments 
	/// SupportByVersion PowerPoint, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744195.aspx </remarks>
	[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), HasIndexProperty(IndexInvoke.Method, "Item")]
 	public class Comments : Collection
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
                    _type = typeof(Comments);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Comments(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Comments(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Comments(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Comments(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Comments(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Comments(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Comments() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Comments(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746408.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Application>(this, "Application", NetOffice.PowerPointApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746446.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16), ProxyResult]
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
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.PowerPointApi.Comment this[Int32 index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Comment>(this, "Item", NetOffice.PowerPointApi.Comment.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745913.aspx </remarks>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="author">string author</param>
		/// <param name="authorInitials">string authorInitials</param>
		/// <param name="text">string text</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Comment Add(Single left, Single top, string author, string authorInitials, string text)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Comment>(this, "Add", NetOffice.PowerPointApi.Comment.LateBindingApiWrapperType, new object[]{ left, top, author, authorInitials, text });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228216.aspx </remarks>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="author">string author</param>
		/// <param name="authorInitials">string authorInitials</param>
		/// <param name="text">string text</param>
		/// <param name="providerID">string providerID</param>
		/// <param name="userID">string userID</param>
		[SupportByVersion("PowerPoint", 15, 16)]
		public NetOffice.PowerPointApi.Comment Add2(Single left, Single top, string author, string authorInitials, string text, string providerID, string userID)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Comment>(this, "Add2", NetOffice.PowerPointApi.Comment.LateBindingApiWrapperType, new object[]{ left, top, author, authorInitials, text, providerID, userID });
		}

		#endregion

		#pragma warning restore
	}
}
