using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface PickerDialog 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860858.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class PickerDialog : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.PickerDialog
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
                    _contractType = typeof(NetOffice.OfficeApi.PickerDialog);
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
                    _type = typeof(PickerDialog);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public PickerDialog() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862371.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual string DataHandlerId
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DataHandlerId");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DataHandlerId", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862526.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual string Title
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Title");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Title", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860248.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual NetOffice.OfficeApi.PickerProperties Properties
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.PickerProperties>(this, "Properties", typeof(NetOffice.OfficeApi.PickerProperties));
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861181.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual NetOffice.OfficeApi.PickerResults CreatePickerResults()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.PickerResults>(this, "CreatePickerResults", typeof(NetOffice.OfficeApi.PickerResults));
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861095.aspx </remarks>
        /// <param name="isMultiSelect">optional bool IsMultiSelect = true</param>
        /// <param name="existingResults">optional NetOffice.OfficeApi.PickerResults ExistingResults = 0</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual NetOffice.OfficeApi.PickerResults Show(object isMultiSelect, object existingResults)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.PickerResults>(this, "Show", typeof(NetOffice.OfficeApi.PickerResults), isMultiSelect, existingResults);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861095.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual NetOffice.OfficeApi.PickerResults Show()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.PickerResults>(this, "Show", typeof(NetOffice.OfficeApi.PickerResults));
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861095.aspx </remarks>
        /// <param name="isMultiSelect">optional bool IsMultiSelect = true</param>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual NetOffice.OfficeApi.PickerResults Show(object isMultiSelect)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.PickerResults>(this, "Show", typeof(NetOffice.OfficeApi.PickerResults), isMultiSelect);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861733.aspx </remarks>
        /// <param name="tokenText">string tokenText</param>
        /// <param name="duplicateDlgMode">Int32 duplicateDlgMode</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual NetOffice.OfficeApi.PickerResults Resolve(string tokenText, Int32 duplicateDlgMode)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.PickerResults>(this, "Resolve", typeof(NetOffice.OfficeApi.PickerResults), tokenText, duplicateDlgMode);
        }

        #endregion

        #pragma warning restore
    }
}
