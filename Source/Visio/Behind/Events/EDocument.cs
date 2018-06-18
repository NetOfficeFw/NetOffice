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
	/// Default implementation of <see cref="NetOffice.VisioApi.EventContracts.EDocument"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class EDocument_SinkHelper : SinkHelper, NetOffice.VisioApi.EventContracts.EDocument
	{
		#region Static
		
		/// <summary>
		/// Interface Id from EDocument
		/// </summary>
		public static readonly string Id = "000D0750-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        /// <exception cref="NetOfficeCOMException">Unexpected error</exception>
        public EDocument_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region EDocument

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

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
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

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
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
		/// <param name="coauthMergeObjects"></param>
        public void AfterDocumentMerge([In, MarshalAs(UnmanagedType.IDispatch)] object coauthMergeObjects)
		{
            if (!Validate("AfterDocumentMerge"))
            {
                Invoker.ReleaseParamsArray(coauthMergeObjects);
                return;
            }

            NetOffice.VisioApi.IVCoauthMergeEvent newcoauthMergeObjects = Factory.CreateEventArgumentObjectFromComProxy(EventClass, coauthMergeObjects) as NetOffice.VisioApi.IVCoauthMergeEvent;
            object[] paramsArray = new object[1];
			paramsArray[0] = newcoauthMergeObjects;
			EventBinding.RaiseCustomEvent("AfterDocumentMerge", ref paramsArray);
		}

		#endregion
	}
	
}
