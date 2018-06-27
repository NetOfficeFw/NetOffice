using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface SignatureSetup 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865226.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class SignatureSetup : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.SignatureSetup
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
                    _contractType = typeof(NetOffice.OfficeApi.SignatureSetup);
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
                    _type = typeof(SignatureSetup);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public SignatureSetup() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860803.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool ReadOnly
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ReadOnly");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863130.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string Id
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Id");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863744.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string SignatureProvider
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "SignatureProvider");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860539.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string SuggestedSigner
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "SuggestedSigner");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SuggestedSigner", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862848.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string SuggestedSignerLine2
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "SuggestedSignerLine2");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SuggestedSignerLine2", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861399.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string SuggestedSignerEmail
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "SuggestedSignerEmail");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SuggestedSignerEmail", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861142.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string SigningInstructions
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "SigningInstructions");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SigningInstructions", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860571.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool AllowComments
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowComments");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowComments", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861064.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool ShowSignDate
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowSignDate");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowSignDate", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864950.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string AdditionalXml
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AdditionalXml");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AdditionalXml", value);
            }
        }

        #endregion

        #region Methods

        #endregion

        #pragma warning restore
    }
}
