using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface DispHTMLAreasCollection 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class DispHTMLAreasCollection : COMObject, NetOffice.MSHTMLApi.DispHTMLAreasCollection
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
                    _contractType = typeof(NetOffice.MSHTMLApi.DispHTMLAreasCollection);
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
                    _type = typeof(DispHTMLAreasCollection);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public DispHTMLAreasCollection() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 length
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "length");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "length", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object _newEnum
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "_newEnum");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 ie8_length
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ie8_length");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object constructor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "constructor");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="index">optional object index</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual object item(object name, object index)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "item", name, index);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual object item()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "item");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual object item(object name)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "item", name);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="tagName">object tagName</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual object tags(object tagName)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "tags", tagName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="element">NetOffice.MSHTMLApi.IHTMLElement element</param>
		/// <param name="before">optional object before</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void add(NetOffice.MSHTMLApi.IHTMLElement element, object before)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "add", element, before);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="element">NetOffice.MSHTMLApi.IHTMLElement element</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void add(NetOffice.MSHTMLApi.IHTMLElement element)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "add", element);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="index">optional Int32 index = -1</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void remove(object index)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "remove", index);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void remove()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "remove");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="urn">object urn</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual object urns(object urn)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "urns", urn);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual object namedItem(string name)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "namedItem", name);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElement2 ie8_item(Int32 index)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElement2>(this, "ie8_item", index);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElement2 ie8_namedItem(string name)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElement2>(this, "ie8_namedItem", name);
		}

		#endregion

		#pragma warning restore
	}
}

