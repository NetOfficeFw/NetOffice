using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface ICTPFactory 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864938.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface ICTPFactory : ICOMObject
    {
        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860563.aspx </remarks>
        /// <param name="cTPAxID">string cTPAxID</param>
        /// <param name="cTPTitle">string cTPTitle</param>
        /// <param name="cTPParentWindow">optional object cTPParentWindow</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.OfficeApi._CustomTaskPane CreateCTP(string cTPAxID, string cTPTitle, object cTPParentWindow);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860563.aspx </remarks>
        /// <param name="cTPAxID">string cTPAxID</param>
        /// <param name="cTPTitle">string cTPTitle</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi._CustomTaskPane CreateCTP(string cTPAxID, string cTPTitle);

        #endregion
    }
}
