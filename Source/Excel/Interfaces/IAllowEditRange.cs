using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
    /// <summary>
    /// Interface IAllowEditRange 
    /// SupportByVersion Excel, 10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public interface IAllowEditRange : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        string Title { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range Range { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.UserAccessList Users { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">string password</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        Int32 ChangePassword(string password);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        Int32 Delete();

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        Int32 Unprotect(object password);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        Int32 Unprotect();

        #endregion
    }
}
