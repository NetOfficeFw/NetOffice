using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface IHTMLImgElement 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IHTMLImgElement : IHTMLControlElement, NetOffice.MSHTMLApi.IHTMLImgElement
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLImgElement);
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
                    _type = typeof(IHTMLImgElement);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLImgElement() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool isMap
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "isMap");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "isMap", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string useMap
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "useMap");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "useMap", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string mimeType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "mimeType");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string fileSize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "fileSize");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string fileCreatedDate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "fileCreatedDate");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string fileModifiedDate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "fileModifiedDate");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string fileUpdatedDate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "fileUpdatedDate");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string protocol
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "protocol");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string href
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "href");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string nameProp
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "nameProp");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object border
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "border");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "border", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 vspace
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "vspace");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "vspace", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 hspace
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "hspace");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "hspace", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string alt
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "alt");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "alt", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string src
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "src");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "src", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string lowsrc
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "lowsrc");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "lowsrc", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string vrml
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "vrml");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "vrml", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string dynsrc
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "dynsrc");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "dynsrc", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string readyState
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "readyState");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool complete
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "complete");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object loop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "loop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "loop", value);
			}
		}
        
		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onload
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onload");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onload", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onerror
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onerror");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onerror", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onabort
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onabort");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onabort", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "name");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "name", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 width
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "width");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "width", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 height
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "height");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "height", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string Start
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Start");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Start", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}

