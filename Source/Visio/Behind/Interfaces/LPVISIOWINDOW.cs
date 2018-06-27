using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi.Behind
{
    /// <summary>
    /// LPVISIOWINDOW
    /// </summary>
    [SyntaxBypass]
    public class LPVISIOWINDOW_ : COMObject, NetOffice.VisioApi.LPVISIOWINDOW_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public LPVISIOWINDOW_() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="reviewerID">optional Int32 reviewerID</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool get_ReviewerMarkupVisible(object reviewerID)
        {
            return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ReviewerMarkupVisible", reviewerID);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="reviewerID">optional Int32 reviewerID</param>
        /// <param name="value">optional bool value</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void set_ReviewerMarkupVisible(object reviewerID, bool value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "ReviewerMarkupVisible", reviewerID, value);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_ReviewerMarkupVisible
        /// </summary>
        /// <param name="reviewerID">optional Int32 reviewerID</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_ReviewerMarkupVisible")]
        public bool ReviewerMarkupVisible(object reviewerID)
        {
            return get_ReviewerMarkupVisible(reviewerID);
        }

        #endregion
    }

    /// <summary>
    /// Interface LPVISIOWINDOW 
    /// SupportByVersion Visio, 11,12,14,15,16
    /// </summary>
    [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class LPVISIOWINDOW : LPVISIOWINDOW_, NetOffice.VisioApi.LPVISIOWINDOW
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
                    _contractType = typeof(NetOffice.VisioApi.LPVISIOWINDOW);
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
                    _type = typeof(LPVISIOWINDOW);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public LPVISIOWINDOW() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVApplication Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 Stat
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Stat");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 ObjectType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ObjectType");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 Type
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Type");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVDocument Document
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "Document");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public NetOffice.VisioApi.IVPage PageAsObj
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVPage>(this, "PageAsObj");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public string PageFromName
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PageFromName");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PageFromName", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Double Zoom
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Zoom");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Zoom", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVSelection Selection
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVSelection>(this, "Selection");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Selection", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 Index
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Index");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 SubType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "SubType");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVEventList EventList
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVEventList>(this, "EventList");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 PersistsEvents
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "PersistsEvents");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int16 WindowHandle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "WindowHandle");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int32 WindowHandle32
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "WindowHandle32");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 ShowRulers
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ShowRulers");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowRulers", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 ShowGrid
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ShowGrid");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowGrid", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 ShowGuides
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ShowGuides");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowGuides", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 ShowConnectPoints
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ShowConnectPoints");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowConnectPoints", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 ShowPageBreaks
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ShowPageBreaks");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowPageBreaks", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public object Page
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Page");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Page", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public object Master
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Master");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 ShowScrollBars
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ShowScrollBars");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowScrollBars", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool Visible
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Visible");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Visible", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string Caption
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Caption");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Caption", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVWindows Windows
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVWindows>(this, "Windows");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int32 WindowState
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "WindowState");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "WindowState", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int32 ViewFit
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ViewFit");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ViewFit", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool IsEditingText
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsEditingText");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool IsEditingOLE
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsEditingOLE");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVWindows Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVWindows>(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVMasterShortcut MasterShortcut
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVMasterShortcut>(this, "MasterShortcut");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int32 ID
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ID");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVWindow ParentWindow
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVWindow>(this, "ParentWindow");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string MergeID
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "MergeID");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MergeID", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string MergeClass
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "MergeClass");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MergeClass", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int32 MergePosition
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "MergePosition");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MergePosition", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool AllowEditing
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowEditing");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowEditing", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Double PageTabWidth
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "PageTabWidth");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PageTabWidth", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool ShowPageTabs
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowPageTabs");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowPageTabs", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool InPlace
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "InPlace");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string MergeCaption
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "MergeCaption");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MergeCaption", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), NativeResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public stdole.Picture Icon
        {
            get
            {
                object[] paramsArray = null;
                object returnItem = Invoker.PropertyGet(this, "Icon", paramsArray);
                return returnItem as stdole.Picture;
            }
            set
            {
                object[] paramsArray = Invoker.ValidateParamsArray(value);
                Invoker.PropertySet(this, "Icon", paramsArray);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVShape Shape
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "Shape");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVCell SelectedCell
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVCell>(this, "SelectedCell");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int32 BackgroundColor
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "BackgroundColor");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BackgroundColor", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int32 BackgroundColorGradient
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "BackgroundColorGradient");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BackgroundColorGradient", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool ShowPageOutline
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowPageOutline");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowPageOutline", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool ScrollLock
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ScrollLock");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ScrollLock", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool ZoomLock
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ZoomLock");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ZoomLock", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public NetOffice.VisioApi.Enums.VisZoomBehavior ZoomBehavior
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisZoomBehavior>(this, "ZoomBehavior");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ZoomBehavior", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult]
        public object[] SelectedMasters
        {
            get
            {
                object[] paramsArray = null;
                object returnItem = Invoker.PropertyGet(this, "SelectedMasters", paramsArray);
                ICOMObject[] newObject = Factory.CreateObjectArrayFromComProxy(this, (object[])returnItem, false);
                return newObject;
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVCharacters SelectedText
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVCharacters>(this, "SelectedText");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "SelectedText", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool ReviewerMarkupVisible
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ReviewerMarkupVisible");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ReviewerMarkupVisible", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVDataRecordset SelectedDataRecordset
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDataRecordset>(this, "SelectedDataRecordset");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "SelectedDataRecordset", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public Int32 SelectedDataRowID
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SelectedDataRowID");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SelectedDataRowID", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVSelection SelectionForDragCopy
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVSelection>(this, "SelectionForDragCopy");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVValidationIssue SelectedValidationIssue
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVValidationIssue>(this, "SelectedValidationIssue");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "SelectedValidationIssue", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Activate()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Activate");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Close()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Close");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void SelectAll()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SelectAll");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void DeselectAll()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DeselectAll");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sheetObject">NetOffice.VisioApi.IVShape sheetObject</param>
        /// <param name="selectAction">Int16 selectAction</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Select(NetOffice.VisioApi.IVShape sheetObject, Int16 selectAction)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Select", sheetObject, selectAction);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Cut()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Cut");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Copy()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Copy");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Paste()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Paste");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Delete()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Duplicate()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Duplicate");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Group()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Group");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Union()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Union");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Combine()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Combine");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Fragment()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Fragment");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void AddToGroup()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AddToGroup");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void RemoveFromGroup()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveFromGroup");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Intersect()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Intersect");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Subtract()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Subtract");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Trim()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Trim");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Join()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Join");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nameArray">String[] nameArray</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void DockedStencils(out String[] nameArray)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
            nameArray = null;
            object[] paramsArray = Invoker.ValidateParamsArray((object)nameArray);
            Invoker.Method(this, "DockedStencils", paramsArray, modifiers);
            nameArray = (String[])paramsArray[0];
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nxFlags">Int32 nxFlags</param>
        /// <param name="nyFlags">Int32 nyFlags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Scroll(Int32 nxFlags, Int32 nyFlags)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Scroll", nxFlags, nyFlags);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void ScrollViewTo(Double x, Double y)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ScrollViewTo", x, y);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pdLeft">Double pdLeft</param>
        /// <param name="pdTop">Double pdTop</param>
        /// <param name="pdWidth">Double pdWidth</param>
        /// <param name="pdHeight">Double pdHeight</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void GetViewRect(out Double pdLeft, out Double pdTop, out Double pdWidth, out Double pdHeight)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true, true, true, true);
            pdLeft = 0;
            pdTop = 0;
            pdWidth = 0;
            pdHeight = 0;
            object[] paramsArray = Invoker.ValidateParamsArray(pdLeft, pdTop, pdWidth, pdHeight);
            Invoker.Method(this, "GetViewRect", paramsArray, modifiers);
            pdLeft = (Double)paramsArray[0];
            pdTop = (Double)paramsArray[1];
            pdWidth = (Double)paramsArray[2];
            pdHeight = (Double)paramsArray[3];
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dLeft">Double dLeft</param>
        /// <param name="dTop">Double dTop</param>
        /// <param name="dWidth">Double dWidth</param>
        /// <param name="dHeight">Double dHeight</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void SetViewRect(Double dLeft, Double dTop, Double dWidth, Double dHeight)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetViewRect", dLeft, dTop, dWidth, dHeight);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pnLeft">Int32 pnLeft</param>
        /// <param name="pnTop">Int32 pnTop</param>
        /// <param name="pnWidth">Int32 pnWidth</param>
        /// <param name="pnHeight">Int32 pnHeight</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void GetWindowRect(out Int32 pnLeft, out Int32 pnTop, out Int32 pnWidth, out Int32 pnHeight)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true, true, true, true);
            pnLeft = 0;
            pnTop = 0;
            pnWidth = 0;
            pnHeight = 0;
            object[] paramsArray = Invoker.ValidateParamsArray(pnLeft, pnTop, pnWidth, pnHeight);
            Invoker.Method(this, "GetWindowRect", paramsArray, modifiers);
            pnLeft = (Int32)paramsArray[0];
            pnTop = (Int32)paramsArray[1];
            pnWidth = (Int32)paramsArray[2];
            pnHeight = (Int32)paramsArray[3];
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nLeft">Int32 nLeft</param>
        /// <param name="nTop">Int32 nTop</param>
        /// <param name="nWidth">Int32 nWidth</param>
        /// <param name="nHeight">Int32 nHeight</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void SetWindowRect(Int32 nLeft, Int32 nTop, Int32 nWidth, Int32 nHeight)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetWindowRect", nLeft, nTop, nWidth, nHeight);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVWindow NewWindow()
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "NewWindow");
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="sheetObject">NetOffice.VisioApi.IVShape sheetObject</param>
        /// <param name="flags">NetOffice.VisioApi.Enums.VisCenterViewFlags flags</param>
        [SupportByVersion("Visio", 14, 15, 16)]
        public void CenterViewOnShape(NetOffice.VisioApi.IVShape sheetObject, NetOffice.VisioApi.Enums.VisCenterViewFlags flags)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CenterViewOnShape", sheetObject, flags);
        }

        #endregion

        #pragma warning restore
    }
}
