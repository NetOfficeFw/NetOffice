using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface ContactCard 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860545.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface ContactCard : _IMsoDispObj
    {
        #region Methods

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863157.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        void Close();

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
        void Show(NetOffice.OfficeApi.Enums.MsoContactCardStyle cardStyle, Int32 rectangleLeft, Int32 rectangleRight, Int32 rectangleTop, Int32 rectangleBottom, Int32 horizontalPosition, object showWithDelay);

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
        void Show(NetOffice.OfficeApi.Enums.MsoContactCardStyle cardStyle, Int32 rectangleLeft, Int32 rectangleRight, Int32 rectangleTop, Int32 rectangleBottom, Int32 horizontalPosition);

        #endregion
    }
}
