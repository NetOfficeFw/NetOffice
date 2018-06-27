using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface DocumentInspector 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862517.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class DocumentInspector : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.DocumentInspector
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
                    _contractType = typeof(NetOffice.OfficeApi.DocumentInspector);
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
                    _type = typeof(DocumentInspector);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public DocumentInspector() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862757.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string Name
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860548.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string Description
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Description");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863644.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861849.aspx </remarks>
        /// <param name="status">NetOffice.OfficeApi.Enums.MsoDocInspectorStatus status</param>
        /// <param name="results">string results</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Inspect(out NetOffice.OfficeApi.Enums.MsoDocInspectorStatus status, out string results)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true, true);
            status = 0;
            results = string.Empty;
            object[] paramsArray = new object[] { status, results };

            InvokerService.InvokeInternal.ExecuteMethodExtended(this, "Inspect", paramsArray, modifiers);

            status = (NetOffice.OfficeApi.Enums.MsoDocInspectorStatus)paramsArray[0];
            results = paramsArray[1] as string;
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863804.aspx </remarks>
        /// <param name="status">NetOffice.OfficeApi.Enums.MsoDocInspectorStatus status</param>
        /// <param name="results">string results</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Fix(out NetOffice.OfficeApi.Enums.MsoDocInspectorStatus status, out string results)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true, true);
            status = 0;
            results = string.Empty;
            object[] paramsArray = new object[] { status, results };

            InvokerService.InvokeInternal.ExecuteMethodExtended(this, "Fix", paramsArray, modifiers);

            status = (NetOffice.OfficeApi.Enums.MsoDocInspectorStatus)paramsArray[0];
            results = paramsArray[1] as string;
        }

        #endregion

        #pragma warning restore
    }
}
