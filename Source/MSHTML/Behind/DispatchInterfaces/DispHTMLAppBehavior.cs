using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface DispHTMLAppBehavior 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class DispHTMLAppBehavior : COMObject, NetOffice.MSHTMLApi.DispHTMLAppBehavior
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
                    _contractType = typeof(NetOffice.MSHTMLApi.DispHTMLAppBehavior);
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
                    _type = typeof(DispHTMLAppBehavior);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public DispHTMLAppBehavior() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string applicationName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "applicationName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "applicationName", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string version
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "version");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "version", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string icon
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "icon");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "icon", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string singleInstance
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "singleInstance");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "singleInstance", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string minimizeButton
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "minimizeButton");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "minimizeButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string maximizeButton
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "maximizeButton");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "maximizeButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string border
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "border");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "border", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string borderStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "borderStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "borderStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string sysMenu
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "sysMenu");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "sysMenu", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string caption
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "caption");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "caption", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string windowState
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "windowState");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "windowState", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string showInTaskBar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "showInTaskBar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "showInTaskBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string commandLine
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "commandLine");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string contextMenu
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "contextMenu");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "contextMenu", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string innerBorder
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "innerBorder");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "innerBorder", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string scroll
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "scroll");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "scroll", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string scrollFlat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "scrollFlat");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "scrollFlat", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string selection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "selection");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "selection", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}

