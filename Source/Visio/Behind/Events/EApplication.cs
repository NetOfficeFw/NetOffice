using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;
using NetOffice.Exceptions;

namespace NetOffice.VisioApi.Behind.EventContracts
{

	/// <summary>
	/// Default implementation of <see cref="NetOffice.VisioApi.EventContracts.EApplication"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class EApplication_SinkHelper : SinkHelper, NetOffice.VisioApi.EventContracts.EApplication
    {
        #region Static

		/// <summary>
		/// Interface Id from EApplication
		/// </summary>
        public static readonly string Id = "000D0B00-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        /// <exception cref="NetOfficeCOMException">Unexpected error</exception>
        public EApplication_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region EApplication Members

        /// <summary>
        /// 
        /// </summary>
        /// <param name="app"></param>
        public void AppActivated([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("AppActivated"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("AppActivated", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="app"></param>
        public void AppDeactivated([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("AppDeactivated"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("AppDeactivated", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="app"></param>
        public void AppObjActivated([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("AppObjActivated"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("AppObjActivated", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="app"></param>
        public void AppObjDeactivated([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("AppObjDeactivated"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("AppObjDeactivated", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="app"></param>
        public void BeforeQuit([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("BeforeQuit"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("BeforeQuit", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="app"></param>
        public void BeforeModal([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("BeforeModal"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("BeforeModal", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="app"></param>
        public void AfterModal([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("AfterModal"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("AfterModal", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="window"></param>
        public void WindowOpened([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("WindowOpened"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("WindowOpened", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="window"></param>
        public void SelectionChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("SelectionChanged"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("SelectionChanged", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="window"></param>
        public void BeforeWindowClosed([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("BeforeWindowClosed"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("BeforeWindowClosed", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="window"></param>
        public void WindowActivated([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("WindowActivated"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("WindowActivated", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="window"></param>
        public void BeforeWindowSelDelete([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("BeforeWindowSelDelete"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("BeforeWindowSelDelete", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="window"></param>
        public void BeforeWindowPageTurn([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("BeforeWindowPageTurn"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("BeforeWindowPageTurn", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="window"></param>
        public void WindowTurnedToPage([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("WindowTurnedToPage"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("WindowTurnedToPage", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
        public void DocumentOpened([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("DocumentOpened"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
            paramsArray[0] = newdoc;
            EventBinding.RaiseCustomEvent("DocumentOpened", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
        public void DocumentCreated([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("DocumentCreated"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument; object[] paramsArray = new object[1];
            paramsArray[0] = newdoc;
            EventBinding.RaiseCustomEvent("DocumentCreated", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
        public void DocumentSaved([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("DocumentSaved"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument; object[] paramsArray = new object[1];
            paramsArray[0] = newdoc;
            EventBinding.RaiseCustomEvent("DocumentSaved", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
        public void DocumentSavedAs([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("DocumentSavedAs"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
            paramsArray[0] = newdoc;
            EventBinding.RaiseCustomEvent("DocumentSavedAs", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
        public void DocumentChanged([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("DocumentChanged"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
            paramsArray[0] = newdoc;
            EventBinding.RaiseCustomEvent("DocumentChanged", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
        public void BeforeDocumentClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("BeforeDocumentClose"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
            paramsArray[0] = newdoc;
            EventBinding.RaiseCustomEvent("BeforeDocumentClose", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="style"></param>
        public void StyleAdded([In, MarshalAs(UnmanagedType.IDispatch)] object style)
        {
            if (!Validate("StyleAdded"))
            {
                Invoker.ReleaseParamsArray(style);
                return;
            }

            NetOffice.VisioApi.IVStyle newStyle = Factory.CreateEventArgumentObjectFromComProxy(EventClass, style) as NetOffice.VisioApi.IVStyle;
            object[] paramsArray = new object[1];
            paramsArray[0] = newStyle;
            EventBinding.RaiseCustomEvent("StyleAdded", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="style"></param>
        public void StyleChanged([In, MarshalAs(UnmanagedType.IDispatch)] object style)
        {
            if (!Validate("StyleChanged"))
            {
                Invoker.ReleaseParamsArray(style);
                return;
            }

            NetOffice.VisioApi.IVStyle newStyle = Factory.CreateEventArgumentObjectFromComProxy(EventClass, style) as NetOffice.VisioApi.IVStyle;
            object[] paramsArray = new object[1];
            paramsArray[0] = newStyle;
            EventBinding.RaiseCustomEvent("StyleChanged", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="style"></param>
        public void BeforeStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style)
        {
            if (!Validate("BeforeStyleDelete"))
            {
                Invoker.ReleaseParamsArray(style);
                return;
            }

            NetOffice.VisioApi.IVStyle newStyle = Factory.CreateEventArgumentObjectFromComProxy(EventClass, style) as NetOffice.VisioApi.IVStyle;
            object[] paramsArray = new object[1];
            paramsArray[0] = newStyle;
            EventBinding.RaiseCustomEvent("BeforeStyleDelete", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="master"></param>
        public void MasterAdded([In, MarshalAs(UnmanagedType.IDispatch)] object master)
        {
            if (!Validate("MasterAdded"))
            {
                Invoker.ReleaseParamsArray(master);
                return;
            }

            NetOffice.VisioApi.IVMaster newMaster = Factory.CreateEventArgumentObjectFromComProxy(EventClass, master) as NetOffice.VisioApi.IVMaster;
            object[] paramsArray = new object[1];
            paramsArray[0] = newMaster;
            EventBinding.RaiseCustomEvent("MasterAdded", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="master"></param>
        public void MasterChanged([In, MarshalAs(UnmanagedType.IDispatch)] object master)
        {
            if (!Validate("MasterChanged"))
            {
                Invoker.ReleaseParamsArray(master);
                return;
            }

            NetOffice.VisioApi.IVMaster newMaster = Factory.CreateEventArgumentObjectFromComProxy(EventClass, master) as NetOffice.VisioApi.IVMaster;
            object[] paramsArray = new object[1];
            paramsArray[0] = newMaster;
            EventBinding.RaiseCustomEvent("MasterChanged", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="master"></param>
        public void BeforeMasterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object master)
        {
            if (!Validate("BeforeMasterDelete"))
            {
                Invoker.ReleaseParamsArray(master);
                return;
            }

            NetOffice.VisioApi.IVMaster newMaster = Factory.CreateEventArgumentObjectFromComProxy(EventClass, master) as NetOffice.VisioApi.IVMaster;
            object[] paramsArray = new object[1];
            paramsArray[0] = newMaster;
            EventBinding.RaiseCustomEvent("BeforeMasterDelete", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="page"></param>
        public void PageAdded([In, MarshalAs(UnmanagedType.IDispatch)] object page)
        {
            if (!Validate("PageAdded"))
            {
                Invoker.ReleaseParamsArray(page);
                return;
            }

            NetOffice.VisioApi.IVPage newPage = Factory.CreateEventArgumentObjectFromComProxy(EventClass, page) as NetOffice.VisioApi.IVPage;
            object[] paramsArray = new object[1];
            paramsArray[0] = newPage;
            EventBinding.RaiseCustomEvent("PageAdded", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="page"></param>
        public void PageChanged([In, MarshalAs(UnmanagedType.IDispatch)] object page)
        {
            if (!Validate("PageChanged"))
            {
                Invoker.ReleaseParamsArray(page);
                return;
            }

            NetOffice.VisioApi.IVPage newPage = Factory.CreateEventArgumentObjectFromComProxy(EventClass, page) as NetOffice.VisioApi.IVPage;
            object[] paramsArray = new object[1];
            paramsArray[0] = newPage;
            EventBinding.RaiseCustomEvent("PageChanged", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="page"></param>
        public void BeforePageDelete([In, MarshalAs(UnmanagedType.IDispatch)] object page)
        {
            if (!Validate("BeforePageDelete"))
            {
                Invoker.ReleaseParamsArray(page);
                return;
            }

            NetOffice.VisioApi.IVPage newPage = Factory.CreateEventArgumentObjectFromComProxy(EventClass, page) as NetOffice.VisioApi.IVPage;
            object[] paramsArray = new object[1];
            paramsArray[0] = newPage;
            EventBinding.RaiseCustomEvent("BeforePageDelete", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shape"></param>
        public void ShapeAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
        {
            if (!Validate("ShapeAdded"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
            paramsArray[0] = newShape;
            EventBinding.RaiseCustomEvent("ShapeAdded", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="selection"></param>
        public void BeforeSelectionDelete([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
        {
            if (!Validate("BeforeSelectionDelete"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSelection;
            EventBinding.RaiseCustomEvent("BeforeSelectionDelete", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shape"></param>
        public void ShapeChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
        {
            if (!Validate("ShapeChanged"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
            paramsArray[0] = newShape;
            EventBinding.RaiseCustomEvent("ShapeChanged", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="selection"></param>
        public void SelectionAdded([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
        {
            if (!Validate("SelectionAdded"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSelection;
            EventBinding.RaiseCustomEvent("SelectionAdded", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shape"></param>
        public void BeforeShapeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
        {
            if (!Validate("BeforeShapeDelete"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
            paramsArray[0] = newShape;
            EventBinding.RaiseCustomEvent("BeforeShapeDelete", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shape"></param>
        public void TextChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
        {
            if (!Validate("TextChanged"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
            paramsArray[0] = newShape;
            EventBinding.RaiseCustomEvent("TextChanged", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cell"></param>
        public void CellChanged([In, MarshalAs(UnmanagedType.IDispatch)] object cell)
        {
            if (!Validate("CellChanged"))
            {
                Invoker.ReleaseParamsArray(cell);
                return;
            }

            NetOffice.VisioApi.IVCell newCell = Factory.CreateEventArgumentObjectFromComProxy(EventClass, cell) as NetOffice.VisioApi.IVCell;
            object[] paramsArray = new object[1];
            paramsArray[0] = newCell;
            EventBinding.RaiseCustomEvent("CellChanged", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="app"></param>
		/// <param name="sequenceNum"></param>
		/// <param name="contextString"></param>
        public void MarkerEvent([In, MarshalAs(UnmanagedType.IDispatch)] object app, [In] object sequenceNum, [In] object contextString)
        {
            if (!Validate("MarkerEvent"))
            {
                Invoker.ReleaseParamsArray(app, sequenceNum, contextString);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            Int32 newSequenceNum = ToInt32(sequenceNum);
            string newContextString = ToString(contextString);
            object[] paramsArray = new object[3];
            paramsArray[0] = newapp;
            paramsArray[1] = newSequenceNum;
            paramsArray[2] = newContextString;
            EventBinding.RaiseCustomEvent("MarkerEvent", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="app"></param>
        public void NoEventsPending([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("NoEventsPending"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("NoEventsPending", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="app"></param>
        public void VisioIsIdle([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("VisioIsIdle"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("VisioIsIdle", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="app"></param>
        public void MustFlushScopeBeginning([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("MustFlushScopeBeginning"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("MustFlushScopeBeginning", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="app"></param>
        public void MustFlushScopeEnded([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("MustFlushScopeEnded"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("MustFlushScopeEnded", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
        public void RunModeEntered([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("RunModeEntered"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
            paramsArray[0] = newdoc;
            EventBinding.RaiseCustomEvent("RunModeEntered", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
        public void DesignModeEntered([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("DesignModeEntered"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
            paramsArray[0] = newdoc;
            EventBinding.RaiseCustomEvent("DesignModeEntered", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
        public void BeforeDocumentSave([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("BeforeDocumentSave"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
            paramsArray[0] = newdoc;
            EventBinding.RaiseCustomEvent("BeforeDocumentSave", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
        public void BeforeDocumentSaveAs([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("BeforeDocumentSaveAs"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
            paramsArray[0] = newdoc;
            EventBinding.RaiseCustomEvent("BeforeDocumentSaveAs", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cell"></param>
        public void FormulaChanged([In, MarshalAs(UnmanagedType.IDispatch)] object cell)
        {
            if (!Validate("FormulaChanged"))
            {
                Invoker.ReleaseParamsArray(cell);
                return;
            }

            NetOffice.VisioApi.IVCell newCell = Factory.CreateEventArgumentObjectFromComProxy(EventClass, cell) as NetOffice.VisioApi.IVCell;
            object[] paramsArray = new object[1];
            paramsArray[0] = newCell;
            EventBinding.RaiseCustomEvent("FormulaChanged", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="connects"></param>
        public void ConnectionsAdded([In, MarshalAs(UnmanagedType.IDispatch)] object connects)
        {
            if (!Validate("ConnectionsAdded"))
            {
                Invoker.ReleaseParamsArray(connects);
                return;
            }

            NetOffice.VisioApi.IVConnects newConnects = Factory.CreateEventArgumentObjectFromComProxy(EventClass, connects) as NetOffice.VisioApi.IVConnects;
            object[] paramsArray = new object[1];
            paramsArray[0] = newConnects;
            EventBinding.RaiseCustomEvent("ConnectionsAdded", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="connects"></param>
        public void ConnectionsDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object connects)
        {
            if (!Validate("ConnectionsDeleted"))
            {
                Invoker.ReleaseParamsArray(connects);
                return;
            }

            NetOffice.VisioApi.IVConnects newConnects = Factory.CreateEventArgumentObjectFromComProxy(EventClass, connects) as NetOffice.VisioApi.IVConnects;
            object[] paramsArray = new object[1];
            paramsArray[0] = newConnects;
            EventBinding.RaiseCustomEvent("ConnectionsDeleted", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="app"></param>
		/// <param name="nScopeID"></param>
		/// <param name="bstrDescription"></param>
        public void EnterScope([In, MarshalAs(UnmanagedType.IDispatch)] object app, [In] object nScopeID, [In] object bstrDescription)
        {
            if (!Validate("EnterScope"))
            {
                Invoker.ReleaseParamsArray(app, nScopeID, bstrDescription);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            Int32 newnScopeID = ToInt32(nScopeID);
            string newbstrDescription = ToString(bstrDescription);
            object[] paramsArray = new object[3];
            paramsArray[0] = newapp;
            paramsArray[1] = newnScopeID;
            paramsArray[2] = newbstrDescription;
            EventBinding.RaiseCustomEvent("EnterScope", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="app"></param>
		/// <param name="nScopeID"></param>
		/// <param name="bstrDescription"></param>
		/// <param name="bErrOrCancelled"></param>
        public void ExitScope([In, MarshalAs(UnmanagedType.IDispatch)] object app, [In] object nScopeID, [In] object bstrDescription, [In] object bErrOrCancelled)
        {
            if (!Validate("ExitScope"))
            {
                Invoker.ReleaseParamsArray(app, nScopeID, bstrDescription, bErrOrCancelled);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            Int32 newnScopeID = ToInt32(nScopeID);
            string newbstrDescription = ToString(bstrDescription);
            bool newbErrOrCancelled = ToBoolean(bErrOrCancelled);
            object[] paramsArray = new object[4];
            paramsArray[0] = newapp;
            paramsArray[1] = newnScopeID;
            paramsArray[2] = newbstrDescription;
            paramsArray[3] = newbErrOrCancelled;
            EventBinding.RaiseCustomEvent("ExitScope", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="app"></param>
        public void QueryCancelQuit([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("QueryCancelQuit"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("QueryCancelQuit", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="app"></param>
        public void QuitCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("QuitCanceled"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("QuitCanceled", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="window"></param>
        public void WindowChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("WindowChanged"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("WindowChanged", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="window"></param>
        public void ViewChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("ViewChanged"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("ViewChanged", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="window"></param>
        public void QueryCancelWindowClose([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("QueryCancelWindowClose"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("QueryCancelWindowClose", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="window"></param>
        public void WindowCloseCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("WindowCloseCanceled"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("WindowCloseCanceled", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
        public void QueryCancelDocumentClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("QueryCancelDocumentClose"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
            paramsArray[0] = newdoc;
            EventBinding.RaiseCustomEvent("QueryCancelDocumentClose", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
        public void DocumentCloseCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("DocumentCloseCanceled"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
            paramsArray[0] = newdoc;
            EventBinding.RaiseCustomEvent("DocumentCloseCanceled", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="style"></param>
        public void QueryCancelStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style)
        {
            if (!Validate("QueryCancelStyleDelete"))
            {
                Invoker.ReleaseParamsArray(style);
                return;
            }

            NetOffice.VisioApi.IVStyle newStyle = Factory.CreateEventArgumentObjectFromComProxy(EventClass, style) as NetOffice.VisioApi.IVStyle;
            object[] paramsArray = new object[1];
            paramsArray[0] = newStyle;
            EventBinding.RaiseCustomEvent("QueryCancelStyleDelete", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="style"></param>
        public void StyleDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object style)
        {
            if (!Validate("StyleDeleteCanceled"))
            {
                Invoker.ReleaseParamsArray(style);
                return;
            }

            NetOffice.VisioApi.IVStyle newStyle = Factory.CreateEventArgumentObjectFromComProxy(EventClass, style) as NetOffice.VisioApi.IVStyle;
            object[] paramsArray = new object[1];
            paramsArray[0] = newStyle;
            EventBinding.RaiseCustomEvent("StyleDeleteCanceled", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="master"></param>
        public void QueryCancelMasterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object master)
        {
            if (!Validate("QueryCancelMasterDelete"))
            {
                Invoker.ReleaseParamsArray(master);
                return;
            }

            NetOffice.VisioApi.IVMaster newMaster = Factory.CreateEventArgumentObjectFromComProxy(EventClass, master) as NetOffice.VisioApi.IVMaster;
            object[] paramsArray = new object[1];
            paramsArray[0] = newMaster;
            EventBinding.RaiseCustomEvent("QueryCancelMasterDelete", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="master"></param>
        public void MasterDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object master)
        {
            if (!Validate("MasterDeleteCanceled"))
            {
                Invoker.ReleaseParamsArray(master);
                return;
            }

            NetOffice.VisioApi.IVMaster newMaster = Factory.CreateEventArgumentObjectFromComProxy(EventClass, master) as NetOffice.VisioApi.IVMaster;
            object[] paramsArray = new object[1];
            paramsArray[0] = newMaster;
            EventBinding.RaiseCustomEvent("MasterDeleteCanceled", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="page"></param>
        public void QueryCancelPageDelete([In, MarshalAs(UnmanagedType.IDispatch)] object page)
        {
            if (!Validate("QueryCancelPageDelete"))
            {
                Invoker.ReleaseParamsArray(page);
                return;
            }

            NetOffice.VisioApi.IVPage newPage = Factory.CreateEventArgumentObjectFromComProxy(EventClass, page) as NetOffice.VisioApi.IVPage;
            object[] paramsArray = new object[1];
            paramsArray[0] = newPage;
            EventBinding.RaiseCustomEvent("QueryCancelPageDelete", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="page"></param>
        public void PageDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object page)
        {
            if (!Validate("PageDeleteCanceled"))
            {
                Invoker.ReleaseParamsArray(page);
                return;
            }

            NetOffice.VisioApi.IVPage newPage = Factory.CreateEventArgumentObjectFromComProxy(EventClass, page) as NetOffice.VisioApi.IVPage;
            object[] paramsArray = new object[1];
            paramsArray[0] = newPage;
            EventBinding.RaiseCustomEvent("PageDeleteCanceled", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shape"></param>
        public void ShapeParentChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
        {
            if (!Validate("ShapeParentChanged"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
            paramsArray[0] = newShape;
            EventBinding.RaiseCustomEvent("ShapeParentChanged", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shape"></param>
        public void BeforeShapeTextEdit([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
        {
            if (!Validate("BeforeShapeTextEdit"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
            paramsArray[0] = newShape;
            EventBinding.RaiseCustomEvent("BeforeShapeTextEdit", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shape"></param>
        public void ShapeExitedTextEdit([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
        {
            if (!Validate("ShapeExitedTextEdit"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
            paramsArray[0] = newShape;
            EventBinding.RaiseCustomEvent("ShapeExitedTextEdit", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="selection"></param>
        public void QueryCancelSelectionDelete([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
        {
            if (!Validate("QueryCancelSelectionDelete"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSelection;
            EventBinding.RaiseCustomEvent("QueryCancelSelectionDelete", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="selection"></param>
        public void SelectionDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
        {
            if (!Validate("SelectionDeleteCanceled"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSelection;
            EventBinding.RaiseCustomEvent("SelectionDeleteCanceled", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="selection"></param>
        public void QueryCancelUngroup([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
        {
            if (!Validate("QueryCancelUngroup"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSelection;
            EventBinding.RaiseCustomEvent("QueryCancelUngroup", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="selection"></param>
        public void UngroupCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
        {
            if (!Validate("UngroupCanceled"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSelection;
            EventBinding.RaiseCustomEvent("UngroupCanceled", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="selection"></param>
        public void QueryCancelConvertToGroup([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
        {
            if (!Validate("QueryCancelConvertToGroup"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSelection;
            EventBinding.RaiseCustomEvent("QueryCancelConvertToGroup", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="selection"></param>
        public void ConvertToGroupCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
        {
            if (!Validate("ConvertToGroupCanceled"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSelection;
            EventBinding.RaiseCustomEvent("ConvertToGroupCanceled", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="app"></param>
        public void QueryCancelSuspend([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("QueryCancelSuspend"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("QueryCancelSuspend", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="app"></param>
        public void SuspendCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("SuspendCanceled"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("SuspendCanceled", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="app"></param>
        public void BeforeSuspend([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("BeforeSuspend"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("BeforeSuspend", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="app"></param>
        public void AfterResume([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("AfterResume"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("AfterResume", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="mSG"></param>
        public void OnKeystrokeMessageForAddon([In, MarshalAs(UnmanagedType.IDispatch)] object mSG)
        {
            if (!Validate("OnKeystrokeMessageForAddon"))
            {
                Invoker.ReleaseParamsArray(mSG);
                return;
            }

            NetOffice.VisioApi.IVMSGWrap newMSG = Factory.CreateEventArgumentObjectFromComProxy(EventClass, mSG) as NetOffice.VisioApi.IVMSGWrap;
            object[] paramsArray = new object[1];
            paramsArray[0] = newMSG;
            EventBinding.RaiseCustomEvent("OnKeystrokeMessageForAddon", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="button"></param>
		/// <param name="keyButtonState"></param>
		/// <param name="x"></param>
		/// <param name="y"></param>
		/// <param name="cancelDefault"></param>
        public void MouseDown([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault)
		{
            if (!Validate("MouseDown"))
            {
                Invoker.ReleaseParamsArray(button, keyButtonState, x, y, cancelDefault);
                return;
            }

			Int32 newButton = ToInt32(button);
			Int32 newKeyButtonState = ToInt32(keyButtonState);
			Double newx = ToDouble(x);
			Double newy = ToDouble(y);
			object[] paramsArray = new object[5];
			paramsArray[0] = newButton;
			paramsArray[1] = newKeyButtonState;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
			paramsArray.SetValue(cancelDefault, 4);
			EventBinding.RaiseCustomEvent("MouseDown", ref paramsArray);

			cancelDefault = ToBoolean(paramsArray[4]);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="button"></param>
		/// <param name="keyButtonState"></param>
		/// <param name="x"></param>
		/// <param name="y"></param>
		/// <param name="cancelDefault"></param>
        public void MouseMove([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault)
		{
            if (!Validate("MouseMove"))
            {
                Invoker.ReleaseParamsArray(button, keyButtonState, x, y, cancelDefault);
                return;
            }

            Int32 newButton = ToInt32(button);
			Int32 newKeyButtonState = ToInt32(keyButtonState);
			Double newx = ToDouble(x);
			Double newy = ToDouble(y);
			object[] paramsArray = new object[5];
			paramsArray[0] = newButton;
			paramsArray[1] = newKeyButtonState;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
			paramsArray.SetValue(cancelDefault, 4);
			EventBinding.RaiseCustomEvent("MouseMove", ref paramsArray);

			cancelDefault = ToBoolean(paramsArray[4]);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="button"></param>
		/// <param name="keyButtonState"></param>
		/// <param name="x"></param>
		/// <param name="y"></param>
		/// <param name="cancelDefault"></param>
        public void MouseUp([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault)
		{
            if (!Validate("MouseUp"))
            {
                Invoker.ReleaseParamsArray(button, keyButtonState, x, y, cancelDefault);
                return;
            }

            Int32 newButton = ToInt32(button);
			Int32 newKeyButtonState = ToInt32(keyButtonState);
			Double newx = ToDouble(x);
			Double newy = ToDouble(y);
			object[] paramsArray = new object[5];
			paramsArray[0] = newButton;
			paramsArray[1] = newKeyButtonState;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
			paramsArray.SetValue(cancelDefault, 4);
			EventBinding.RaiseCustomEvent("MouseUp", ref paramsArray);

			cancelDefault = ToBoolean(paramsArray[4]);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="keyCode"></param>
		/// <param name="keyButtonState"></param>
		/// <param name="cancelDefault"></param>
        public void KeyDown([In] object keyCode, [In] object keyButtonState, [In] [Out] ref object cancelDefault)
        {
            if (!Validate("KeyDown"))
            {
                Invoker.ReleaseParamsArray(keyCode, keyButtonState, cancelDefault);
                return;
            }

			Int32 newKeyCode = ToInt32(keyCode);
			Int32 newKeyButtonState = ToInt32(keyButtonState);
			object[] paramsArray = new object[3];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newKeyButtonState;
			paramsArray.SetValue(cancelDefault, 2);
			EventBinding.RaiseCustomEvent("KeyDown", ref paramsArray);

			cancelDefault = ToBoolean(paramsArray[2]);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="keyAscii"></param>
		/// <param name="cancelDefault"></param>
        public void KeyPress([In] object keyAscii, [In] [Out] ref object cancelDefault)
		{
            if (!Validate("KeyPress"))
            {
                Invoker.ReleaseParamsArray(keyAscii, cancelDefault);
                return;
            }

			Int32 newKeyAscii = ToInt32(keyAscii);
			object[] paramsArray = new object[2];
			paramsArray[0] = newKeyAscii;
			paramsArray.SetValue(cancelDefault, 1);
			EventBinding.RaiseCustomEvent("KeyPress", ref paramsArray);

			cancelDefault = ToBoolean(paramsArray[1]);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="keyCode"></param>
		/// <param name="keyButtonState"></param>
		/// <param name="cancelDefault"></param>
        public void KeyUp([In] object keyCode, [In] object keyButtonState, [In] [Out] ref object cancelDefault)
        {
            if (!Validate("KeyUp"))
            {
                Invoker.ReleaseParamsArray(keyCode, keyButtonState, cancelDefault);
                return;
            }

			Int32 newKeyCode = ToInt32(keyCode);
			Int32 newKeyButtonState = ToInt32(keyButtonState);
			object[] paramsArray = new object[3];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newKeyButtonState;
			paramsArray.SetValue(cancelDefault, 2);
			EventBinding.RaiseCustomEvent("KeyUp", ref paramsArray);

			cancelDefault = ToBoolean(paramsArray[2]);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="app"></param>
        public void QueryCancelSuspendEvents([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("QueryCancelSuspendEvents"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			EventBinding.RaiseCustomEvent("QueryCancelSuspendEvents", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="app"></param>
        public void SuspendEventsCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
            if (!Validate("SuspendEventsCanceled"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			EventBinding.RaiseCustomEvent("SuspendEventsCanceled", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="app"></param>
        public void BeforeSuspendEvents([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("BeforeSuspendEvents"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			EventBinding.RaiseCustomEvent("BeforeSuspendEvents", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="app"></param>
        public void AfterResumeEvents([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
            if (!Validate("AfterResumeEvents"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			EventBinding.RaiseCustomEvent("AfterResumeEvents", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="selection"></param>
        public void QueryCancelGroup([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
        {
            if (!Validate("QueryCancelGroup"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
			paramsArray[0] = newSelection;
			EventBinding.RaiseCustomEvent("QueryCancelGroup", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="selection"></param>
        public void GroupCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
            if (!Validate("GroupCanceled"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
			paramsArray[0] = newSelection;
			EventBinding.RaiseCustomEvent("GroupCanceled", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shape"></param>
        public void ShapeDataGraphicChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
		{
            if (!Validate("ShapeDataGraphicChanged"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
			paramsArray[0] = newShape;
			EventBinding.RaiseCustomEvent("ShapeDataGraphicChanged", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="dataRecordset"></param>
        public void BeforeDataRecordsetDelete([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordset)
        {
            if (!Validate("BeforeDataRecordsetDelete"))
            {
                Invoker.ReleaseParamsArray(dataRecordset);
                return;
            }

            NetOffice.VisioApi.IVDataRecordset newDataRecordset = Factory.CreateEventArgumentObjectFromComProxy(EventClass, dataRecordset) as NetOffice.VisioApi.IVDataRecordset;
            object[] paramsArray = new object[1];
			paramsArray[0] = newDataRecordset;
			EventBinding.RaiseCustomEvent("BeforeDataRecordsetDelete", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="dataRecordsetChanged"></param>
        public void DataRecordsetChanged([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordsetChanged)
		{
            if (!Validate("BeforeDataRecordsetDelete"))
            {
                Invoker.ReleaseParamsArray(dataRecordsetChanged);
                return;
            }

            NetOffice.VisioApi.IVDataRecordsetChangedEvent newDataRecordsetChanged = Factory.CreateEventArgumentObjectFromComProxy(EventClass, dataRecordsetChanged) as NetOffice.VisioApi.IVDataRecordsetChangedEvent;
            object[] paramsArray = new object[1];
			paramsArray[0] = newDataRecordsetChanged;
			EventBinding.RaiseCustomEvent("DataRecordsetChanged", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="dataRecordset"></param>
        public void DataRecordsetAdded([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordset)
		{
            if (!Validate("DataRecordsetAdded"))
            {
                Invoker.ReleaseParamsArray(dataRecordset);
                return;
            }

            NetOffice.VisioApi.IVDataRecordset newDataRecordset = Factory.CreateEventArgumentObjectFromComProxy(EventClass, dataRecordset) as NetOffice.VisioApi.IVDataRecordset;
            object[] paramsArray = new object[1];
			paramsArray[0] = newDataRecordset;
			EventBinding.RaiseCustomEvent("DataRecordsetAdded", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shape"></param>
		/// <param name="dataRecordsetID"></param>
		/// <param name="dataRowID"></param>
        public void ShapeLinkAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shape, [In] object dataRecordsetID, [In] object dataRowID)
        {
            if (!Validate("ShapeLinkAdded"))
            {
                Invoker.ReleaseParamsArray(shape, dataRecordsetID, dataRowID);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            Int32 newDataRecordsetID = ToInt32(dataRecordsetID);
			Int32 newDataRowID = ToInt32(dataRowID);
			object[] paramsArray = new object[3];
			paramsArray[0] = newShape;
			paramsArray[1] = newDataRecordsetID;
			paramsArray[2] = newDataRowID;
			EventBinding.RaiseCustomEvent("ShapeLinkAdded", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shape"></param>
		/// <param name="dataRecordsetID"></param>
		/// <param name="dataRowID"></param>
        public void ShapeLinkDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object shape, [In] object dataRecordsetID, [In] object dataRowID)
        {
            if (!Validate("ShapeLinkDeleted"))
            {
                Invoker.ReleaseParamsArray(shape, dataRecordsetID, dataRowID);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            Int32 newDataRecordsetID = ToInt32(dataRecordsetID);
			Int32 newDataRowID = ToInt32(dataRowID);
			object[] paramsArray = new object[3];
			paramsArray[0] = newShape;
			paramsArray[1] = newDataRecordsetID;
			paramsArray[2] = newDataRowID;
			EventBinding.RaiseCustomEvent("ShapeLinkDeleted", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
        public void AfterRemoveHiddenInformation([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("AfterRemoveHiddenInformation"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			EventBinding.RaiseCustomEvent("AfterRemoveHiddenInformation", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shapePair"></param>
        public void ContainerRelationshipAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair)
		{
            if (!Validate("ContainerRelationshipAdded"))
            {
                Invoker.ReleaseParamsArray(shapePair);
                return;
            }

            NetOffice.VisioApi.IVRelatedShapePairEvent newShapePair = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shapePair) as NetOffice.VisioApi.IVRelatedShapePairEvent;
            object[] paramsArray = new object[1];
			paramsArray[0] = newShapePair;
			EventBinding.RaiseCustomEvent("ContainerRelationshipAdded", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shapePair"></param>
        public void ContainerRelationshipDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair)
		{
            if (!Validate("ContainerRelationshipDeleted"))
            {
                Invoker.ReleaseParamsArray(shapePair);
                return;
            }

            NetOffice.VisioApi.IVRelatedShapePairEvent newShapePair = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shapePair) as NetOffice.VisioApi.IVRelatedShapePairEvent;
            object[] paramsArray = new object[1];
			paramsArray[0] = newShapePair;
			EventBinding.RaiseCustomEvent("ContainerRelationshipDeleted", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shapePair"></param>
        public void CalloutRelationshipAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair)
		{
            if (!Validate("CalloutRelationshipAdded"))
            {
                Invoker.ReleaseParamsArray(shapePair);
                return;
            }

            NetOffice.VisioApi.IVRelatedShapePairEvent newShapePair = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shapePair) as NetOffice.VisioApi.IVRelatedShapePairEvent;
            object[] paramsArray = new object[1];
			paramsArray[0] = newShapePair;
			EventBinding.RaiseCustomEvent("CalloutRelationshipAdded", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shapePair"></param>
        public void CalloutRelationshipDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair)
		{
            if (!Validate("CalloutRelationshipDeleted"))
            {
                Invoker.ReleaseParamsArray(shapePair);
                return;
            }

            NetOffice.VisioApi.IVRelatedShapePairEvent newShapePair = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shapePair) as NetOffice.VisioApi.IVRelatedShapePairEvent;
            object[] paramsArray = new object[1];
			paramsArray[0] = newShapePair;
			EventBinding.RaiseCustomEvent("CalloutRelationshipDeleted", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="ruleSet"></param>
        public void RuleSetValidated([In, MarshalAs(UnmanagedType.IDispatch)] object ruleSet)
		{
            if (!Validate("RuleSetValidated"))
            {
                Invoker.ReleaseParamsArray(ruleSet);
                return;
            }

            NetOffice.VisioApi.IVValidationRuleSet newRuleSet = Factory.CreateEventArgumentObjectFromComProxy(EventClass, ruleSet) as NetOffice.VisioApi.IVValidationRuleSet;
            object[] paramsArray = new object[1];
			paramsArray[0] = newRuleSet;
			EventBinding.RaiseCustomEvent("RuleSetValidated", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="replaceShapes"></param>
        public void QueryCancelReplaceShapes([In, MarshalAs(UnmanagedType.IDispatch)] object replaceShapes)
		{
            if (!Validate("QueryCancelReplaceShapes"))
            {
                Invoker.ReleaseParamsArray(replaceShapes);
                return;
            }

            NetOffice.VisioApi.IVReplaceShapesEvent newreplaceShapes = Factory.CreateEventArgumentObjectFromComProxy(EventClass, replaceShapes) as NetOffice.VisioApi.IVReplaceShapesEvent;
            object[] paramsArray = new object[1];
			paramsArray[0] = newreplaceShapes;
			EventBinding.RaiseCustomEvent("QueryCancelReplaceShapes", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="replaceShapes"></param>
        public void ReplaceShapesCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object replaceShapes)
		{
            if (!Validate("ReplaceShapesCanceled"))
            {
                Invoker.ReleaseParamsArray(replaceShapes);
                return;
            }

            NetOffice.VisioApi.IVReplaceShapesEvent newreplaceShapes = Factory.CreateEventArgumentObjectFromComProxy(EventClass, replaceShapes) as NetOffice.VisioApi.IVReplaceShapesEvent;
            object[] paramsArray = new object[1];
			paramsArray[0] = newreplaceShapes;
			EventBinding.RaiseCustomEvent("ReplaceShapesCanceled", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="replaceShapes"></param>
        public void BeforeReplaceShapes([In, MarshalAs(UnmanagedType.IDispatch)] object replaceShapes)
		{
            if (!Validate("BeforeReplaceShapes"))
            {
                Invoker.ReleaseParamsArray(replaceShapes);
                return;
            }

            NetOffice.VisioApi.IVReplaceShapesEvent newreplaceShapes = Factory.CreateEventArgumentObjectFromComProxy(EventClass, replaceShapes) as NetOffice.VisioApi.IVReplaceShapesEvent;
            object[] paramsArray = new object[1];
			paramsArray[0] = newreplaceShapes;
			EventBinding.RaiseCustomEvent("BeforeReplaceShapes", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="sel"></param>
        public void AfterReplaceShapes([In, MarshalAs(UnmanagedType.IDispatch)] object sel)
        {
            if (!Validate("AfterReplaceShapes"))
            {
                Invoker.ReleaseParamsArray(sel);
                return;
            }

            NetOffice.VisioApi.IVSelection newsel = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sel) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
			paramsArray[0] = newsel;
			EventBinding.RaiseCustomEvent("AfterReplaceShapes", ref paramsArray);
		}

		#endregion
	}
	
}
