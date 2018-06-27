using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.WordApi;

namespace NetOffice.WordApi.Behind
{
	/// <summary>
	/// DispatchInterface TextInput 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838939.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class TextInput : COMObject, NetOffice.WordApi.TextInput
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
                    _contractType = typeof(NetOffice.WordApi.TextInput);
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
                    _type = typeof(TextInput);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public TextInput() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193885.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", typeof(NetOffice.WordApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192550.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Creator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838708.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838485.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool Valid
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Valid");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821318.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string Default
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Default");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Default", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197119.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdTextFormFieldType Type
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdTextFormFieldType>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196338.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string Format
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Format");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834845.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Width
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Width");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Width", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822366.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Clear()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Clear");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840130.aspx </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdTextFormFieldType type</param>
		/// <param name="_default">optional object default</param>
		/// <param name="format">optional object format</param>
		/// <param name="enabled">optional object enabled</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void EditType(NetOffice.WordApi.Enums.WdTextFormFieldType type, object _default, object format, object enabled)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EditType", type, _default, format, enabled);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840130.aspx </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdTextFormFieldType type</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void EditType(NetOffice.WordApi.Enums.WdTextFormFieldType type)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EditType", type);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840130.aspx </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdTextFormFieldType type</param>
		/// <param name="_default">optional object default</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void EditType(NetOffice.WordApi.Enums.WdTextFormFieldType type, object _default)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EditType", type, _default);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840130.aspx </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdTextFormFieldType type</param>
		/// <param name="_default">optional object default</param>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void EditType(NetOffice.WordApi.Enums.WdTextFormFieldType type, object _default, object format)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EditType", type, _default, format);
		}

		#endregion

		#pragma warning restore
	}
}


