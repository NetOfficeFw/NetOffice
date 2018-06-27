using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OutlookApi;

namespace NetOffice.OutlookApi.Behind
{
	/// <summary>
	/// DispatchInterface _PropertyAccessor 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _PropertyAccessor : COMObject, NetOffice.OutlookApi._PropertyAccessor
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
                    _contractType = typeof(NetOffice.OutlookApi._PropertyAccessor);
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
                    _type = typeof(_PropertyAccessor);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _PropertyAccessor() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865030.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Application>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869821.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual NetOffice.OutlookApi.Enums.OlObjectClass Class
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlObjectClass>(this, "Class");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869435.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._NameSpace Session
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._NameSpace>(this, "Session");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866730.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868350.aspx </remarks>
		/// <param name="schemaName">string schemaName</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual object GetProperty(string schemaName)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GetProperty", schemaName);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862751.aspx </remarks>
		/// <param name="schemaName">string schemaName</param>
		/// <param name="value">object value</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual void SetProperty(string schemaName, object value)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetProperty", schemaName, value);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869865.aspx </remarks>
		/// <param name="schemaNames">object schemaNames</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual object GetProperties(object schemaNames)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GetProperties", schemaNames);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868862.aspx </remarks>
		/// <param name="schemaNames">object schemaNames</param>
		/// <param name="values">object values</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual object SetProperties(object schemaNames, object values)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SetProperties", schemaNames, values);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868342.aspx </remarks>
		/// <param name="value">DateTime value</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual DateTime UTCToLocalTime(DateTime value)
		{
			return InvokerService.InvokeInternal.ExecuteDateTimeMethodGet(this, "UTCToLocalTime", value);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868909.aspx </remarks>
		/// <param name="value">DateTime value</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual DateTime LocalTimeToUTC(DateTime value)
		{
			return InvokerService.InvokeInternal.ExecuteDateTimeMethodGet(this, "LocalTimeToUTC", value);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862123.aspx </remarks>
		/// <param name="value">string value</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual object StringToBinary(string value)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "StringToBinary", value);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864468.aspx </remarks>
		/// <param name="value">object value</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string BinaryToString(object value)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "BinaryToString", value);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868076.aspx </remarks>
		/// <param name="schemaName">string schemaName</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual void DeleteProperty(string schemaName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteProperty", schemaName);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869707.aspx </remarks>
		/// <param name="schemaNames">object schemaNames</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual object DeleteProperties(object schemaNames)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DeleteProperties", schemaNames);
		}

		#endregion

		#pragma warning restore
	}
}

