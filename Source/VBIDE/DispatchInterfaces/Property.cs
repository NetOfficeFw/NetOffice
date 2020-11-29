using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi
{
	/// <summary>
	/// DispatchInterface Property 
	/// SupportByVersion VBIDE, 12,14,5.3
	/// </summary>
	[SupportByVersion("VBIDE", 12,14,5.3)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Property : COMObject
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
                    _type = typeof(Property);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Property(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Property(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Property(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Property(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Property(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Property(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Property() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Property(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get/Set
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public object Value
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Value");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Value", value);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get/Set
		/// </summary>
		/// <param name="index1">object index1</param>
		/// <param name="index2">optional object index2</param>
		/// <param name="index3">optional object index3</param>
		/// <param name="index4">optional object index4</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_IndexedValue(object index1, object index2, object index3, object index4)
		{
			return Factory.ExecuteVariantPropertyGet(this, "IndexedValue", index1, index2, index3, index4);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get/Set
		/// </summary>
		/// <param name="index1">object index1</param>
		/// <param name="index2">optional object index2</param>
		/// <param name="index3">optional object index3</param>
		/// <param name="index4">optional object index4</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_IndexedValue(object index1, object index2, object index3, object index4, object value)
		{
			Factory.ExecutePropertySet(this, "IndexedValue", new object[]{ index1, index2, index3, index4, value });
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Alias for get_IndexedValue
		/// </summary>
		/// <param name="index1">object index1</param>
		/// <param name="index2">optional object index2</param>
		/// <param name="index3">optional object index3</param>
		/// <param name="index4">optional object index4</param>
		[SupportByVersion("VBIDE", 12,14,5.3), Redirect("get_IndexedValue")]
		public object IndexedValue(object index1, object index2, object index3, object index4)
		{
			return get_IndexedValue(index1, index2, index3, index4);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get/Set
		/// </summary>
		/// <param name="index1">object index1</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_IndexedValue(object index1)
		{
			return Factory.ExecuteVariantPropertyGet(this, "IndexedValue", index1);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get/Set
		/// </summary>
		/// <param name="index1">object index1</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_IndexedValue(object index1, object value)
		{
			Factory.ExecutePropertySet(this, "IndexedValue", index1, value);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Alias for get_IndexedValue
		/// </summary>
		/// <param name="index1">object index1</param>
		[SupportByVersion("VBIDE", 12,14,5.3), Redirect("get_IndexedValue")]
		public object IndexedValue(object index1)
		{
			return get_IndexedValue(index1);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get/Set
		/// </summary>
		/// <param name="index1">object index1</param>
		/// <param name="index2">optional object index2</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_IndexedValue(object index1, object index2)
		{
			return Factory.ExecuteVariantPropertyGet(this, "IndexedValue", index1, index2);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get/Set
		/// </summary>
		/// <param name="index1">object index1</param>
		/// <param name="index2">optional object index2</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_IndexedValue(object index1, object index2, object value)
		{
			Factory.ExecutePropertySet(this, "IndexedValue", index1, index2, value);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Alias for get_IndexedValue
		/// </summary>
		/// <param name="index1">object index1</param>
		/// <param name="index2">optional object index2</param>
		[SupportByVersion("VBIDE", 12,14,5.3), Redirect("get_IndexedValue")]
		public object IndexedValue(object index1, object index2)
		{
			return get_IndexedValue(index1, index2);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get/Set
		/// </summary>
		/// <param name="index1">object index1</param>
		/// <param name="index2">optional object index2</param>
		/// <param name="index3">optional object index3</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_IndexedValue(object index1, object index2, object index3)
		{
			return Factory.ExecuteVariantPropertyGet(this, "IndexedValue", index1, index2, index3);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get/Set
		/// </summary>
		/// <param name="index1">object index1</param>
		/// <param name="index2">optional object index2</param>
		/// <param name="index3">optional object index3</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_IndexedValue(object index1, object index2, object index3, object value)
		{
			Factory.ExecutePropertySet(this, "IndexedValue", index1, index2, index3, value);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Alias for get_IndexedValue
		/// </summary>
		/// <param name="index1">object index1</param>
		/// <param name="index2">optional object index2</param>
		/// <param name="index3">optional object index3</param>
		[SupportByVersion("VBIDE", 12,14,5.3), Redirect("get_IndexedValue")]
		public object IndexedValue(object index1, object index2, object index3)
		{
			return get_IndexedValue(index1, index2, index3);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public Int16 NumIndices
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "NumIndices");
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		[BaseResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VBIDEApi.Application Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VBIDEApi.Application>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VBIDEApi.Properties Parent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.Properties>(this, "Parent", NetOffice.VBIDEApi.Properties.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.VBE VBE
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBE>(this, "VBE", NetOffice.VBIDEApi.VBE.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.Properties Collection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.Properties>(this, "Collection", NetOffice.VBIDEApi.Properties.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3), ProxyResult]
		public object Object
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Object");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Object", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
