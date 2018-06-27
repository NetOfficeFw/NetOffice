using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface MsoDebugOptions 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class MsoDebugOptions : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.MsoDebugOptions
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
                    _contractType = typeof(NetOffice.OfficeApi.MsoDebugOptions);
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
                    _type = typeof(MsoDebugOptions);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public MsoDebugOptions() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 FeatureReports
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "FeatureReports");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FeatureReports", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual bool OutputToDebugger
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "OutputToDebugger");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OutputToDebugger", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual bool OutputToFile
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "OutputToFile");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OutputToFile", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual bool OutputToMessageBox
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "OutputToMessageBox");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OutputToMessageBox", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        public virtual object UnitTestManager
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "UnitTestManager");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 15,16
        /// </summary>
        /// <param name="bstrTagToIgnore">string bstrTagToIgnore</param>
        [SupportByVersion("Office", 15, 16)]
        public virtual void AddIgnoredAssertTag(string bstrTagToIgnore)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AddIgnoredAssertTag", bstrTagToIgnore);
        }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// </summary>
        /// <param name="bstrTagToIgnore">string bstrTagToIgnore</param>
        [SupportByVersion("Office", 15, 16)]
        public virtual void RemoveIgnoredAssertTag(string bstrTagToIgnore)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveIgnoredAssertTag", bstrTagToIgnore);
        }

        #endregion

        #pragma warning restore
    }
}
