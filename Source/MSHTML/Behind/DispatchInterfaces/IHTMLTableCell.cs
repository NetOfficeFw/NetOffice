using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface IHTMLTableCell 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IHTMLTableCell : COMObject, NetOffice.MSHTMLApi.IHTMLTableCell
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLTableCell);
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
                    _type = typeof(IHTMLTableCell);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLTableCell() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 rowSpan
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "rowSpan");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "rowSpan", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 colSpan
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "colSpan");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "colSpan", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string align
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "align");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "align", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string vAlign
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "vAlign");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "vAlign", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object bgColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "bgColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "bgColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool noWrap
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "noWrap");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "noWrap", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string background
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "background");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "background", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object borderColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "borderColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "borderColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object borderColorLight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "borderColorLight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "borderColorLight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object borderColorDark
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "borderColorDark");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "borderColorDark", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object width
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "width");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "width", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object height
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "height");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "height", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 cellIndex
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "cellIndex");
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}

