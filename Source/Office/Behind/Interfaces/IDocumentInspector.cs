using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// Interface IDocumentInspector 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861808.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IDocumentInspector : COMObject, NetOffice.OfficeApi.IDocumentInspector
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
                    _contractType = typeof(NetOffice.OfficeApi.IDocumentInspector);
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
                    _type = typeof(IDocumentInspector);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IDocumentInspector() : base()
		{

		}

        #endregion

        #region Properties

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862465.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="desc">string desc</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 GetInfo(out string name, out string desc)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true, true);
            name = string.Empty;
            desc = string.Empty;
            object[] paramsArray = new object[] { name, desc };

            Int32 returnItem = InvokerService.InvokeInternal.ExecuteInt32MethodGetExtended(this, "GetInfo", paramsArray, modifiers);

            name = paramsArray[0] as string;
            desc = paramsArray[1] as string;
            return returnItem;
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861133.aspx </remarks>
        /// <param name="doc">object doc</param>
        /// <param name="status">NetOffice.OfficeApi.Enums.MsoDocInspectorStatus status</param>
        /// <param name="result">string result</param>
        /// <param name="action">string action</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 Inspect(object doc, out NetOffice.OfficeApi.Enums.MsoDocInspectorStatus status, out string result, out string action)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, true, true, true);
            status = 0;
            result = string.Empty;
            action = string.Empty;
            object[] paramsArray = new object[] { doc, status, result, action };

            Int32 returnItem = InvokerService.InvokeInternal.ExecuteInt32MethodGetExtended(this, "Inspect", paramsArray, modifiers);

            status = (NetOffice.OfficeApi.Enums.MsoDocInspectorStatus)paramsArray[1];
            result = paramsArray[2] as string;
            action = paramsArray[3] as string;
            return returnItem;
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864114.aspx </remarks>
        /// <param name="doc">object doc</param>
        /// <param name="hwnd">Int32 hwnd</param>
        /// <param name="status">NetOffice.OfficeApi.Enums.MsoDocInspectorStatus status</param>
        /// <param name="result">string result</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 Fix(object doc, Int32 hwnd, out NetOffice.OfficeApi.Enums.MsoDocInspectorStatus status, out string result)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, false, true, true);
            status = 0;
            result = string.Empty;
            object[] paramsArray = new object[] { doc, hwnd, status, result };

            Int32 returnItem = InvokerService.InvokeInternal.ExecuteInt32MethodGetExtended(this, "Fix", paramsArray, modifiers);

            status = (NetOffice.OfficeApi.Enums.MsoDocInspectorStatus)paramsArray[2];
            result = paramsArray[3] as string;
            return returnItem;
        }

        #endregion

        #pragma warning restore
    }
}
