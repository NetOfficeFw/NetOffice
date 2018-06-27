using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface LanguageSettings 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863125.aspx </remarks>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class LanguageSettings : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.LanguageSettings
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
                    _contractType = typeof(NetOffice.OfficeApi.LanguageSettings);
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
                    _type = typeof(LanguageSettings);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public LanguageSettings() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863438.aspx </remarks>
        /// <param name="id">NetOffice.OfficeApi.Enums.MsoAppLanguageID id</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 get_LanguageID(NetOffice.OfficeApi.Enums.MsoAppLanguageID id)
        {
            return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "LanguageID", id);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_LanguageID
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863438.aspx </remarks>
        /// <param name="id">NetOffice.OfficeApi.Enums.MsoAppLanguageID id</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_LanguageID")]
        public virtual Int32 LanguageID(NetOffice.OfficeApi.Enums.MsoAppLanguageID id)
        {
            return get_LanguageID(id);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861143.aspx </remarks>
        /// <param name="lid">NetOffice.OfficeApi.Enums.MsoLanguageID lid</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool get_LanguagePreferredForEditing(NetOffice.OfficeApi.Enums.MsoLanguageID lid)
        {
            return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "LanguagePreferredForEditing", lid);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_LanguagePreferredForEditing
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861143.aspx </remarks>
        /// <param name="lid">NetOffice.OfficeApi.Enums.MsoLanguageID lid</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_LanguagePreferredForEditing")]
        public virtual bool LanguagePreferredForEditing(NetOffice.OfficeApi.Enums.MsoLanguageID lid)
        {
            return get_LanguagePreferredForEditing(lid);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862054.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        #endregion

        #region Methods

        #endregion

        #pragma warning restore
    }
}
