using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface SharedWorkspaceTasks 
    /// SupportByVersion Office, 11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864958.aspx </remarks>
    [SupportByVersion("Office", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Office", 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("000C037A-0000-0000-C000-000000000046")]
    public interface SharedWorkspaceTasks : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.SharedWorkspaceTask>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.OfficeApi.SharedWorkspaceTask this[Int32 index] { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862401.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        Int32 Count { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862065.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861502.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        bool ItemCountExceeded { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865453.aspx </remarks>
        /// <param name="title">string title</param>
        /// <param name="status">optional object status</param>
        /// <param name="priority">optional object priority</param>
        /// <param name="assignee">optional object assignee</param>
        /// <param name="description">optional object description</param>
        /// <param name="dueDate">optional object dueDate</param>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.SharedWorkspaceTask Add(string title, object status, object priority, object assignee, object description, object dueDate);

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865453.aspx </remarks>
        /// <param name="title">string title</param>
        [CustomMethod]
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.SharedWorkspaceTask Add(string title);

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865453.aspx </remarks>
        /// <param name="title">string title</param>
        /// <param name="status">optional object status</param>
        [CustomMethod]
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.SharedWorkspaceTask Add(string title, object status);

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865453.aspx </remarks>
        /// <param name="title">string title</param>
        /// <param name="status">optional object status</param>
        /// <param name="priority">optional object priority</param>
        [CustomMethod]
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.SharedWorkspaceTask Add(string title, object status, object priority);

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865453.aspx </remarks>
        /// <param name="title">string title</param>
        /// <param name="status">optional object status</param>
        /// <param name="priority">optional object priority</param>
        /// <param name="assignee">optional object assignee</param>
        [CustomMethod]
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.SharedWorkspaceTask Add(string title, object status, object priority, object assignee);

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865453.aspx </remarks>
        /// <param name="title">string title</param>
        /// <param name="status">optional object status</param>
        /// <param name="priority">optional object priority</param>
        /// <param name="assignee">optional object assignee</param>
        /// <param name="description">optional object description</param>
        [CustomMethod]
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.SharedWorkspaceTask Add(string title, object status, object priority, object assignee, object description);

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.SharedWorkspaceTask>

        /// <summary>
        /// SupportByVersion Office, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.OfficeApi.SharedWorkspaceTask> GetEnumerator();

        #endregion
    }
}
