using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// XMLNamespace
	/// </summary>
	[SyntaxBypass]
 	public class XMLNamespace_ : COMObject
	{
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public XMLNamespace_(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public XMLNamespace_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNamespace_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNamespace_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNamespace_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNamespace_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNamespace_() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNamespace_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="allUsers">optional bool allUsers</param>
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNamespace.Location"/>
		[SupportByVersion("Word", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Location(object allUsers)
		{
			return Factory.ExecuteStringPropertyGet(this, "Location", allUsers);
		}

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="allUsers">optional bool allUsers</param>
        /// <param name="value">optional string value</param>
        [SupportByVersion("Word", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Location(object allUsers, string value)
		{
			Factory.ExecutePropertySet(this, "Location", value, allUsers);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Alias for get_Location
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNamespace.Location"/> </remarks>
		/// <param name="allUsers">optional bool allUsers</param>
		[SupportByVersion("Word", 11,12,14,15,16), Redirect("get_Location")]
		public string Location(object allUsers)
		{
			return get_Location(allUsers);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="allUsers">optional bool allUsers</param>
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNamespace.Alias"/>
		[SupportByVersion("Word", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Alias(object allUsers)
		{
			return Factory.ExecuteStringPropertyGet(this, "Alias", allUsers);
		}

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="allUsers">optional bool allUsers</param>
        /// <param name="value">optional string value</param>
        [SupportByVersion("Word", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Alias(object allUsers, string value)
		{
			Factory.ExecutePropertySet(this, "Alias", value, allUsers);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Alias for get_Alias
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNamespace.Alias"/> </remarks>
		/// <param name="allUsers">optional bool allUsers</param>
		[SupportByVersion("Word", 11,12,14,15,16), Redirect("get_Alias")]
		public string Alias(object allUsers)
		{
			return get_Alias(allUsers);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="allUsers">optional bool allUsers</param>
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNamespace.DefaultTransform"/>
		[SupportByVersion("Word", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.XSLTransform get_DefaultTransform(object allUsers)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XSLTransform>(this, "DefaultTransform", NetOffice.WordApi.XSLTransform.LateBindingApiWrapperType, allUsers);
		}

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="allUsers">optional bool allUsers</param>
        /// <param name="value">optional XSLTransform value</param>
        [SupportByVersion("Word", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_DefaultTransform(object allUsers, NetOffice.WordApi.XSLTransform value)
		{
			Factory.ExecutePropertySet(this, "DefaultTransform", value, allUsers);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Alias for get_DefaultTransform
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNamespace.DefaultTransform"/> </remarks>
		/// <param name="allUsers">optional bool allUsers</param>
		[SupportByVersion("Word", 11,12,14,15,16), Redirect("get_DefaultTransform")]
		public NetOffice.WordApi.XSLTransform DefaultTransform(object allUsers)
		{
			return get_DefaultTransform(allUsers);
		}

		#endregion

		#region Methods

		#endregion
	}

	/// <summary>
	/// DispatchInterface XMLNamespace 
	/// SupportByVersion Word, 11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNamespace"/> </remarks>
	[SupportByVersion("Word", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class XMLNamespace : XMLNamespace_
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
                    _type = typeof(XMLNamespace);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public XMLNamespace(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public XMLNamespace(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNamespace(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNamespace(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNamespace(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNamespace(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNamespace() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNamespace(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNamespace.Application"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", NetOffice.WordApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNamespace.Creator"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNamespace.Parent"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.URI"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public string URI
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "URI");
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNamespace.Location"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public string Location
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Location");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Location", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNamespace.Alias"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public string Alias
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Alias");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Alias", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNamespace.XSLTransforms"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XSLTransforms XSLTransforms
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XSLTransforms>(this, "XSLTransforms", NetOffice.WordApi.XSLTransforms.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNamespace.DefaultTransform"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XSLTransform DefaultTransform
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XSLTransform>(this, "DefaultTransform", NetOffice.WordApi.XSLTransform.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "DefaultTransform", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNamespace.AttachToDocument"/> </remarks>
		/// <param name="document">object document</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void AttachToDocument(object document)
		{
			 Factory.ExecuteMethod(this, "AttachToDocument", document);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNamespace.Delete"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		#endregion

		#pragma warning restore
	}
}
