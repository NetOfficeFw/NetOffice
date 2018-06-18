using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
    /// <summary>
    /// LPVISIOWINDOW
    /// </summary>
    [SyntaxBypass]
    public interface LPVISIOWINDOW_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="reviewerID">optional Int32 reviewerID</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        bool get_ReviewerMarkupVisible(object reviewerID);


        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="reviewerID">optional Int32 reviewerID</param>
        /// <param name="value">optional bool value</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_ReviewerMarkupVisible(object reviewerID, bool value);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_ReviewerMarkupVisible
        /// </summary>
        /// <param name="reviewerID">optional Int32 reviewerID</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_ReviewerMarkupVisible")]
        bool ReviewerMarkupVisible(object reviewerID);

        #endregion
    }
    
    /// <summary>
    /// Interface LPVISIOWINDOW 
    /// SupportByVersion Visio, 11,12,14,15,16
    /// </summary>
    [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
	[TypeId("00000000-0000-0000-0000-000000000000")]
    public interface LPVISIOWINDOW : LPVISIOWINDOW_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVApplication Application { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 Stat { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 ObjectType { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 Type { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVDocument Document { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.VisioApi.IVPage PageAsObj { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string PageFromName { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Double Zoom { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVSelection Selection { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 Index { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 SubType { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVEventList EventList { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 PersistsEvents { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int16 WindowHandle { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int32 WindowHandle32 { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 ShowRulers { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 ShowGrid { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 ShowGuides { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 ShowConnectPoints { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 ShowPageBreaks { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        object Page { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        object Master { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 ShowScrollBars { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool Visible { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string Caption { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVWindows Windows { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int32 WindowState { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int32 ViewFit { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool IsEditingText { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool IsEditingOLE { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVWindows Parent { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVMasterShortcut MasterShortcut { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int32 ID { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVWindow ParentWindow { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string MergeID { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string MergeClass { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int32 MergePosition { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool AllowEditing { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Double PageTabWidth { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool ShowPageTabs { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool InPlace { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string MergeCaption { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), NativeResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        stdole.Picture Icon { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVShape Shape { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVCell SelectedCell { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int32 BackgroundColor { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int32 BackgroundColorGradient { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool ShowPageOutline { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool ScrollLock { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool ZoomLock { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        NetOffice.VisioApi.Enums.VisZoomBehavior ZoomBehavior { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult]
        object[] SelectedMasters { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVCharacters SelectedText { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        new bool ReviewerMarkupVisible { get; set; }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVDataRecordset SelectedDataRecordset { get; set; }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        Int32 SelectedDataRowID { get; set; }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVSelection SelectionForDragCopy { get; }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVValidationIssue SelectedValidationIssue { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Activate();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Close();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void SelectAll();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void DeselectAll();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sheetObject">NetOffice.VisioApi.IVShape sheetObject</param>
        /// <param name="selectAction">Int16 selectAction</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Select(NetOffice.VisioApi.IVShape sheetObject, Int16 selectAction);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Cut();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Copy();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Paste();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Delete();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Duplicate();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Group();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Union();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Combine();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Fragment();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void AddToGroup();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void RemoveFromGroup();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Intersect();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Subtract();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Trim();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Join();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nameArray">String[] nameArray</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void DockedStencils(out String[] nameArray);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nxFlags">Int32 nxFlags</param>
        /// <param name="nyFlags">Int32 nyFlags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Scroll(Int32 nxFlags, Int32 nyFlags);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void ScrollViewTo(Double x, Double y);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pdLeft">Double pdLeft</param>
        /// <param name="pdTop">Double pdTop</param>
        /// <param name="pdWidth">Double pdWidth</param>
        /// <param name="pdHeight">Double pdHeight</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void GetViewRect(out Double pdLeft, out Double pdTop, out Double pdWidth, out Double pdHeight);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dLeft">Double dLeft</param>
        /// <param name="dTop">Double dTop</param>
        /// <param name="dWidth">Double dWidth</param>
        /// <param name="dHeight">Double dHeight</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void SetViewRect(Double dLeft, Double dTop, Double dWidth, Double dHeight);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pnLeft">Int32 pnLeft</param>
        /// <param name="pnTop">Int32 pnTop</param>
        /// <param name="pnWidth">Int32 pnWidth</param>
        /// <param name="pnHeight">Int32 pnHeight</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void GetWindowRect(out Int32 pnLeft, out Int32 pnTop, out Int32 pnWidth, out Int32 pnHeight);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nLeft">Int32 nLeft</param>
        /// <param name="nTop">Int32 nTop</param>
        /// <param name="nWidth">Int32 nWidth</param>
        /// <param name="nHeight">Int32 nHeight</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void SetWindowRect(Int32 nLeft, Int32 nTop, Int32 nWidth, Int32 nHeight);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVWindow NewWindow();

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="sheetObject">NetOffice.VisioApi.IVShape sheetObject</param>
        /// <param name="flags">NetOffice.VisioApi.Enums.VisCenterViewFlags flags</param>
        [SupportByVersion("Visio", 14, 15, 16)]
        void CenterViewOnShape(NetOffice.VisioApi.IVShape sheetObject, NetOffice.VisioApi.Enums.VisCenterViewFlags flags);

        #endregion
    }
}
