using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface _CommandBars 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Office", 9, 10, 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("000C0302-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OfficeApi.CommandBars))]
    public interface _CommandBars : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.CommandBar>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862425.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.OfficeApi.CommandBarControl ActionControl { get; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863075.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.CommandBar ActiveMenuBar { get; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860520.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Count { get; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863160.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        bool DisplayTooltips { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864956.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        bool DisplayKeysInTooltips { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.OfficeApi.CommandBar this[object index] { get; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864068.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        bool LargeButtons { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864076.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoMenuAnimation MenuAnimationStyle { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862543.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="ids">Int32 ids</param>
        /// <param name="pbstrName">string pbstrName</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int32 get_IdsString(Int32 ids, out string pbstrName);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_IdsString
        /// </summary>
        /// <param name="ids">Int32 ids</param>
        /// <param name="pbstrName">string pbstrName</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_IdsString")]
        Int32 IdsString(Int32 ids, out string pbstrName);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="tmc">Int32 tmc</param>
        /// <param name="pbstrName">string pbstrName</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int32 get_TmcGetName(Int32 tmc, out string pbstrName);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_TmcGetName
        /// </summary>
        /// <param name="tmc">Int32 tmc</param>
        /// <param name="pbstrName">string pbstrName</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_TmcGetName")]
        Int32 TmcGetName(Int32 tmc, out string pbstrName);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860590.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        bool AdaptiveMenus { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860823.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        bool DisplayFonts { get; set; }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864631.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        bool DisableCustomize { get; set; }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863405.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        bool DisableAskAQuestionDropdown { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861773.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="position">optional object position</param>
        /// <param name="menuBar">optional object menuBar</param>
        /// <param name="temporary">optional object temporary</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.CommandBar Add(object name, object position, object menuBar, object temporary);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861773.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.CommandBar Add();

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861773.aspx </remarks>
        /// <param name="name">optional object name</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.CommandBar Add(object name);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861773.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="position">optional object position</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.CommandBar Add(object name, object position);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861773.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="position">optional object position</param>
        /// <param name="menuBar">optional object menuBar</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.CommandBar Add(object name, object position, object menuBar);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860267.aspx </remarks>
        /// <param name="type">optional object type</param>
        /// <param name="id">optional object id</param>
        /// <param name="tag">optional object tag</param>
        /// <param name="visible">optional object visible</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.OfficeApi.CommandBarControl FindControl(object type, object id, object tag, object visible);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860267.aspx </remarks>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.CommandBarControl FindControl();

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860267.aspx </remarks>
        /// <param name="type">optional object type</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.CommandBarControl FindControl(object type);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860267.aspx </remarks>
        /// <param name="type">optional object type</param>
        /// <param name="id">optional object id</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.CommandBarControl FindControl(object type, object id);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860267.aspx </remarks>
        /// <param name="type">optional object type</param>
        /// <param name="id">optional object id</param>
        /// <param name="tag">optional object tag</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.CommandBarControl FindControl(object type, object id, object tag);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861062.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void ReleaseFocus();

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862556.aspx </remarks>
        /// <param name="type">optional object type</param>
        /// <param name="id">optional object id</param>
        /// <param name="tag">optional object tag</param>
        /// <param name="visible">optional object visible</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.CommandBarControls FindControls(object type, object id, object tag, object visible);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862556.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.CommandBarControls FindControls();

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862556.aspx </remarks>
        /// <param name="type">optional object type</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.CommandBarControls FindControls(object type);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862556.aspx </remarks>
        /// <param name="type">optional object type</param>
        /// <param name="id">optional object id</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.CommandBarControls FindControls(object type, object id);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862556.aspx </remarks>
        /// <param name="type">optional object type</param>
        /// <param name="id">optional object id</param>
        /// <param name="tag">optional object tag</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.CommandBarControls FindControls(object type, object id, object tag);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="tbidOrName">optional object tbidOrName</param>
        /// <param name="position">optional object position</param>
        /// <param name="menuBar">optional object menuBar</param>
        /// <param name="temporary">optional object temporary</param>
        /// <param name="tbtrProtection">optional object tbtrProtection</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.CommandBar AddEx(object tbidOrName, object position, object menuBar, object temporary, object tbtrProtection);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.CommandBar AddEx();

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="tbidOrName">optional object tbidOrName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.CommandBar AddEx(object tbidOrName);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="tbidOrName">optional object tbidOrName</param>
        /// <param name="position">optional object position</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.CommandBar AddEx(object tbidOrName, object position);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="tbidOrName">optional object tbidOrName</param>
        /// <param name="position">optional object position</param>
        /// <param name="menuBar">optional object menuBar</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.CommandBar AddEx(object tbidOrName, object position, object menuBar);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="tbidOrName">optional object tbidOrName</param>
        /// <param name="position">optional object position</param>
        /// <param name="menuBar">optional object menuBar</param>
        /// <param name="temporary">optional object temporary</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.CommandBar AddEx(object tbidOrName, object position, object menuBar, object temporary);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862419.aspx </remarks>
        /// <param name="idMso">string idMso</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ExecuteMso(string idMso);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862202.aspx </remarks>
        /// <param name="idMso">string idMso</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool GetEnabledMso(string idMso);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863712.aspx </remarks>
        /// <param name="idMso">string idMso</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool GetVisibleMso(string idMso);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863149.aspx </remarks>
        /// <param name="idMso">string idMso</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool GetPressedMso(string idMso);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860585.aspx </remarks>
        /// <param name="idMso">string idMso</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        string GetLabelMso(string idMso);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860790.aspx </remarks>
        /// <param name="idMso">string idMso</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        string GetScreentipMso(string idMso);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864975.aspx </remarks>
        /// <param name="idMso">string idMso</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        string GetSupertipMso(string idMso);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861156.aspx </remarks>
        /// <param name="idMso">string idMso</param>
        /// <param name="width">Int32 width</param>
        /// <param name="height">Int32 height</param>
        [SupportByVersion("Office", 12, 14, 15, 16), NativeResult]
        stdole.Picture GetImageMso(string idMso, Int32 width, Int32 height);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863478.aspx </remarks>
        /// <param name="hwnd">Int32 hwnd</param>
        [SupportByVersion("Office", 14, 15, 16)]
        void CommitRenderingTransaction(Int32 hwnd);

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.CommandBar>

        /// <summary>
        /// SupportByVersion Office, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.OfficeApi.CommandBar> GetEnumerator();

        #endregion
    }
}
