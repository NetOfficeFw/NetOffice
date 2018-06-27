using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface ContactCard 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860545.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class ContactCard : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.ContactCard
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
                    _contractType = typeof(NetOffice.OfficeApi.ContactCard);
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
                    _type = typeof(ContactCard);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ContactCard() : base()
		{

		}

		#endregion

        #region Properties

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863157.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual void Close()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Close");
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861819.aspx </remarks>
        /// <param name="cardStyle">NetOffice.OfficeApi.Enums.MsoContactCardStyle cardStyle</param>
        /// <param name="rectangleLeft">Int32 rectangleLeft</param>
        /// <param name="rectangleRight">Int32 rectangleRight</param>
        /// <param name="rectangleTop">Int32 rectangleTop</param>
        /// <param name="rectangleBottom">Int32 rectangleBottom</param>
        /// <param name="horizontalPosition">Int32 horizontalPosition</param>
        /// <param name="showWithDelay">optional bool ShowWithDelay = false</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual void Show(NetOffice.OfficeApi.Enums.MsoContactCardStyle cardStyle, Int32 rectangleLeft, Int32 rectangleRight, Int32 rectangleTop, Int32 rectangleBottom, Int32 horizontalPosition, object showWithDelay)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Show", new object[] { cardStyle, rectangleLeft, rectangleRight, rectangleTop, rectangleBottom, horizontalPosition, showWithDelay });
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861819.aspx </remarks>
        /// <param name="cardStyle">NetOffice.OfficeApi.Enums.MsoContactCardStyle cardStyle</param>
        /// <param name="rectangleLeft">Int32 rectangleLeft</param>
        /// <param name="rectangleRight">Int32 rectangleRight</param>
        /// <param name="rectangleTop">Int32 rectangleTop</param>
        /// <param name="rectangleBottom">Int32 rectangleBottom</param>
        /// <param name="horizontalPosition">Int32 horizontalPosition</param>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual void Show(NetOffice.OfficeApi.Enums.MsoContactCardStyle cardStyle, Int32 rectangleLeft, Int32 rectangleRight, Int32 rectangleTop, Int32 rectangleBottom, Int32 horizontalPosition)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Show", new object[] { cardStyle, rectangleLeft, rectangleRight, rectangleTop, rectangleBottom, horizontalPosition });
        }

        #endregion

        #pragma warning restore
    }
}
