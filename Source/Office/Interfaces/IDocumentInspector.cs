using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// Interface IDocumentInspector 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861808.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public interface IDocumentInspector : ICOMObject
    {
        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862465.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="desc">string desc</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 GetInfo(out string name, out string desc);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861133.aspx </remarks>
        /// <param name="doc">object doc</param>
        /// <param name="status">NetOffice.OfficeApi.Enums.MsoDocInspectorStatus status</param>
        /// <param name="result">string result</param>
        /// <param name="action">string action</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 Inspect(object doc, out NetOffice.OfficeApi.Enums.MsoDocInspectorStatus status, out string result, out string action);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864114.aspx </remarks>
        /// <param name="doc">object doc</param>
        /// <param name="hwnd">Int32 hwnd</param>
        /// <param name="status">NetOffice.OfficeApi.Enums.MsoDocInspectorStatus status</param>
        /// <param name="result">string result</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 Fix(object doc, Int32 hwnd, out NetOffice.OfficeApi.Enums.MsoDocInspectorStatus status, out string result);

        #endregion
    }
}
