using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OutlookApi;

namespace NetOffice.OutlookApi.Behind
{
	/// <summary>
	/// DispatchInterface _Rule 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Rule : COMObject, NetOffice.OutlookApi._Rule
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
                    _contractType = typeof(NetOffice.OutlookApi._Rule);
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
                    _type = typeof(_Rule);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _Rule() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867534.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868179.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866703.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869390.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866230.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860647.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual Int32 ExecutionOrder
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ExecutionOrder");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ExecutionOrder", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866197.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual NetOffice.OutlookApi.Enums.OlRuleType RuleType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlRuleType>(this, "RuleType");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868097.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual bool Enabled
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Enabled");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Enabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864224.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual bool IsLocalRule
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsLocalRule");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862765.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual NetOffice.OutlookApi.RuleActions Actions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.RuleActions>(this, "Actions", typeof(NetOffice.OutlookApi.RuleActions));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869579.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual NetOffice.OutlookApi.RuleConditions Conditions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.RuleConditions>(this, "Conditions", typeof(NetOffice.OutlookApi.RuleConditions));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867239.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual NetOffice.OutlookApi.RuleConditions Exceptions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.RuleConditions>(this, "Exceptions", typeof(NetOffice.OutlookApi.RuleConditions));
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864433.aspx </remarks>
		/// <param name="showProgress">optional object showProgress</param>
		/// <param name="folder">optional object folder</param>
		/// <param name="includeSubfolders">optional object includeSubfolders</param>
		/// <param name="ruleExecuteOption">optional object ruleExecuteOption</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual void Execute(object showProgress, object folder, object includeSubfolders, object ruleExecuteOption)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Execute", showProgress, folder, includeSubfolders, ruleExecuteOption);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864433.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual void Execute()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Execute");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864433.aspx </remarks>
		/// <param name="showProgress">optional object showProgress</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual void Execute(object showProgress)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Execute", showProgress);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864433.aspx </remarks>
		/// <param name="showProgress">optional object showProgress</param>
		/// <param name="folder">optional object folder</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual void Execute(object showProgress, object folder)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Execute", showProgress, folder);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864433.aspx </remarks>
		/// <param name="showProgress">optional object showProgress</param>
		/// <param name="folder">optional object folder</param>
		/// <param name="includeSubfolders">optional object includeSubfolders</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual void Execute(object showProgress, object folder, object includeSubfolders)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Execute", showProgress, folder, includeSubfolders);
		}

		#endregion

		#pragma warning restore
	}
}


