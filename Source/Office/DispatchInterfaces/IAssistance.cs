using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface IAssistance 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864589.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface IAssistance : ICOMObject
    {
        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860570.aspx </remarks>
        /// <param name="helpId">optional string HelpId = </param>
        /// <param name="scope">optional string Scope = </param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ShowHelp(object helpId, object scope);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860570.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ShowHelp();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860570.aspx </remarks>
        /// <param name="helpId">optional string HelpId = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ShowHelp(object helpId);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862805.aspx </remarks>
        /// <param name="query">string query</param>
        /// <param name="scope">optional string Scope = </param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void SearchHelp(string query, object scope);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862805.aspx </remarks>
        /// <param name="query">string query</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void SearchHelp(string query);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861230.aspx </remarks>
        /// <param name="helpId">string helpId</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void SetDefaultContext(string helpId);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865260.aspx </remarks>
        /// <param name="helpId">string helpId</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ClearDefaultContext(string helpId);

        #endregion
    }
}
