using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.WordApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Word", 11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("00020A01-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ApplicationEvents4
	{
		[SupportByVersion("Word", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void Startup();

		[SupportByVersion("Word", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void Quit();

		[SupportByVersion("Word", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void DocumentChange();

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void DocumentOpen([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void DocumentBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void DocumentBeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("saveAsUI", SinkArgumentType.Bool)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void DocumentBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object saveAsUI, [In] [Out] ref object cancel);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
		void NewDocument([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("wn", typeof(WordApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10)]
		void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("wn", typeof(WordApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(11)]
		void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("sel", typeof(WordApi.Selection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12)]
		void WindowSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sel);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("sel", typeof(WordApi.Selection))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
		void WindowBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("sel", typeof(WordApi.Selection))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(14)]
		void WindowBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(15)]
		void EPostagePropertyDialog([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16)]
		void EPostageInsert([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("docResult", typeof(WordApi.Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(17)]
		void MailMergeAfterMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object docResult);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(18)]
		void MailMergeAfterRecordMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("startRecord", SinkArgumentType.Int32)]
        [SinkArgument("endRecord", SinkArgumentType.Int32)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(19)]
		void MailMergeBeforeMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object startRecord, [In] object endRecord, [In] [Out] ref object cancel);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(20)]
		void MailMergeBeforeRecordMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(21)]
		void MailMergeDataSourceLoad([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("handled", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(22)]
		void MailMergeDataSourceValidate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object handled);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(23)]
		void MailMergeWizardSendToCustom([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("fromState", SinkArgumentType.Int32)]
        [SinkArgument("toState", SinkArgumentType.Int32)]
        [SinkArgument("handled", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(24)]
		void MailMergeWizardStateChange([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object fromState, [In] [Out] ref object toState, [In] [Out] ref object handled);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("wn", typeof(WordApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(25)]
		void WindowSize([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("sel", typeof(WordApi.Selection))]
        [SinkArgument("oldXMLNode", typeof(WordApi.XMLNode))]
        [SinkArgument("newXMLNode", typeof(WordApi.XMLNode))]
        [SinkArgument("reason", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(26)]
		void XMLSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In, MarshalAs(UnmanagedType.IDispatch)] object oldXMLNode, [In, MarshalAs(UnmanagedType.IDispatch)] object newXMLNode, [In] [Out] ref object reason);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("xMLNode", typeof(WordApi.XMLNode))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(27)]
		void XMLValidationError([In, MarshalAs(UnmanagedType.IDispatch)] object xMLNode);

        [SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("syncEventType", SinkArgumentType.Enum, typeof(OfficeApi.Enums.MsoSyncEventType))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(28)]
		void DocumentSync([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object syncEventType);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("cpDeliveryAddrStart", SinkArgumentType.Int32)]
        [SinkArgument("cpDeliveryAddrEnd", SinkArgumentType.Int32)]
        [SinkArgument("cpReturnAddrStart", SinkArgumentType.Int32)]
        [SinkArgument("cpReturnAddrEnd", SinkArgumentType.Int32)]
        [SinkArgument("xaWidth", SinkArgumentType.Int32)]
        [SinkArgument("yaHeight", SinkArgumentType.Int32)]
        [SinkArgument("bstrPrinterName", SinkArgumentType.String)]
        [SinkArgument("bstrPaperFeed", SinkArgumentType.String)]
        [SinkArgument("fPrint", SinkArgumentType.Bool)]
        [SinkArgument("fCancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(29)]
		void EPostageInsertEx([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object cpDeliveryAddrStart, [In] object cpDeliveryAddrEnd, [In] object cpReturnAddrStart, [In] object cpReturnAddrEnd, [In] object xaWidth, [In] object yaHeight, [In] object bstrPrinterName, [In] object bstrPaperFeed, [In] object fPrint, [In] [Out] ref object fCancel);

		[SupportByVersion("Word", 12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("handled", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(30)]
		void MailMergeDataSourceValidate2([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object handled);

		[SupportByVersion("Word", 14,15,16)]
        [SinkArgument("pvWindow", typeof(WordApi.ProtectedViewWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(31)]
		void ProtectedViewWindowOpen([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow);

		[SupportByVersion("Word", 14,15,16)]
        [SinkArgument("pvWindow", typeof(WordApi.ProtectedViewWindow))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32)]
		void ProtectedViewWindowBeforeEdit([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow, [In] [Out] ref object cancel);

		[SupportByVersion("Word", 14,15,16)]
        [SinkArgument("pvWindow", typeof(WordApi.ProtectedViewWindow))]
        [SinkArgument("newCloseReason", SinkArgumentType.Int32)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(33)]
		void ProtectedViewWindowBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow, [In] object closeReason, [In] [Out] ref object cancel);

		[SupportByVersion("Word", 14,15,16)]
        [SinkArgument("pvWindow", typeof(WordApi.ProtectedViewWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(34)]
		void ProtectedViewWindowSize([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow);

		[SupportByVersion("Word", 14,15,16)]
        [SinkArgument("pvWindow", typeof(WordApi.ProtectedViewWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(35)]
		void ProtectedViewWindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow);

		[SupportByVersion("Word", 14,15,16)]
        [SinkArgument("pvWindow", typeof(WordApi.ProtectedViewWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(36)]
		void ProtectedViewWindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ApplicationEvents4_SinkHelper : SinkHelper, ApplicationEvents4
	{
		#region Static
		
		public static readonly string Id = "00020A01-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

		public ApplicationEvents4_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}

        #endregion

        #region ApplicationEvents4

        public void Startup()
        {
            if (!Validate("Startup"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("Startup", ref paramsArray);
        }

        public void Quit()
        {
            if (!Validate("Quit"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("Quit", ref paramsArray);
        }

        public void DocumentChange()
        {
            if (!Validate("DocumentChange"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("DocumentChange", ref paramsArray);
        }

        public void DocumentOpen([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("DocumentOpen"))
            {
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, NetOffice.WordApi.Document.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
            paramsArray[0] = newDoc;
            EventBinding.RaiseCustomEvent("DocumentOpen", ref paramsArray);
        }

        public void DocumentBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel)
        {
            if (!Validate("DocumentBeforeClose"))
            {
                Invoker.ReleaseParamsArray(doc, cancel);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, NetOffice.WordApi.Document.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newDoc;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("DocumentBeforeClose", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        public void DocumentBeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel)
        {
            if (!Validate("DocumentBeforePrint"))
            {
                Invoker.ReleaseParamsArray(doc, cancel);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, NetOffice.WordApi.Document.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newDoc;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("DocumentBeforePrint", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        public void DocumentBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object saveAsUI, [In] [Out] ref object cancel)
        {
            if (!Validate("DocumentBeforeSave"))
            {
                Invoker.ReleaseParamsArray(doc, saveAsUI, cancel);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, NetOffice.WordApi.Document.LateBindingApiWrapperType);
            object[] paramsArray = new object[3];
            paramsArray[0] = newDoc;
            paramsArray.SetValue(saveAsUI, 1);
            paramsArray.SetValue(cancel, 2);
            EventBinding.RaiseCustomEvent("DocumentBeforeSave", ref paramsArray);

            saveAsUI = ToBoolean(paramsArray[1]);
            cancel = ToBoolean(paramsArray[2]);
        }

        public void NewDocument([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("NewDocument"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, NetOffice.WordApi.Document.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
            paramsArray[0] = newDoc;
            EventBinding.RaiseCustomEvent("NewDocument", ref paramsArray);
        }

        public void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("WindowActivate"))
            {
                Invoker.ReleaseParamsArray(doc, wn);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, NetOffice.WordApi.Document.LateBindingApiWrapperType);
            NetOffice.WordApi.Window newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Window>(EventClass, wn, NetOffice.WordApi.Window.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newDoc;
            paramsArray[1] = newWn;
            EventBinding.RaiseCustomEvent("WindowActivate", ref paramsArray);
        }

        public void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("WindowDeactivate"))
            {
                Invoker.ReleaseParamsArray(doc, wn);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, NetOffice.WordApi.Document.LateBindingApiWrapperType);
            NetOffice.WordApi.Window newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Window>(EventClass, wn, NetOffice.WordApi.Window.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newDoc;
            paramsArray[1] = newWn;
            EventBinding.RaiseCustomEvent("WindowDeactivate", ref paramsArray);
        }

        public void WindowSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sel)
        {
            if (!Validate("WindowSelectionChange"))
            {
                Invoker.ReleaseParamsArray(sel);
                return;
            }

            NetOffice.WordApi.Selection newSel = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Selection>(EventClass, sel, NetOffice.WordApi.Selection.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
            paramsArray[0] = newSel;
            EventBinding.RaiseCustomEvent("WindowSelectionChange", ref paramsArray);
        }

        public void WindowBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel)
        {
            if (!Validate("WindowBeforeRightClick"))
            {
                Invoker.ReleaseParamsArray(sel, cancel);
                return;
            }

            NetOffice.WordApi.Selection newSel = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Selection>(EventClass, sel, NetOffice.WordApi.Selection.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newSel;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("WindowBeforeRightClick", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        public void WindowBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel)
        {
            if (!Validate("WindowBeforeDoubleClick"))
            {
                Invoker.ReleaseParamsArray(sel, cancel);
                return;
            }

            NetOffice.WordApi.Selection newSel = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Selection>(EventClass, sel, NetOffice.WordApi.Selection.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newSel;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("WindowBeforeDoubleClick", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        public void EPostagePropertyDialog([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("EPostagePropertyDialog"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, NetOffice.WordApi.Document.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
            paramsArray[0] = newDoc;
            EventBinding.RaiseCustomEvent("EPostagePropertyDialog", ref paramsArray);
        }

        public void EPostageInsert([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("EPostageInsert"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, NetOffice.WordApi.Document.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
            paramsArray[0] = newDoc;
            EventBinding.RaiseCustomEvent("EPostageInsert", ref paramsArray);
        }

        public void MailMergeAfterMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object docResult)
        {
            if (!Validate("MailMergeAfterMerge"))
            {
                Invoker.ReleaseParamsArray(doc, docResult);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, NetOffice.WordApi.Document.LateBindingApiWrapperType);
            NetOffice.WordApi.Document newDocResult = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, docResult, NetOffice.WordApi.Document.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newDoc;
            paramsArray[1] = newDocResult;
            EventBinding.RaiseCustomEvent("MailMergeAfterMerge", ref paramsArray);
        }

        public void MailMergeAfterRecordMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("MailMergeAfterRecordMerge"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, NetOffice.WordApi.Document.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
            paramsArray[0] = newDoc;
            EventBinding.RaiseCustomEvent("MailMergeAfterRecordMerge", ref paramsArray);
        }

        public void MailMergeBeforeMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object startRecord, [In] object endRecord, [In] [Out] ref object cancel)
        {
            if (!Validate("MailMergeBeforeMerge"))
            {
                Invoker.ReleaseParamsArray(doc, startRecord, endRecord, cancel);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, NetOffice.WordApi.Document.LateBindingApiWrapperType);
            Int32 newStartRecord = ToInt32(startRecord);
            Int32 newEndRecord = ToInt32(endRecord);
            object[] paramsArray = new object[4];
            paramsArray[0] = newDoc;
            paramsArray[1] = newStartRecord;
            paramsArray[2] = newEndRecord;
            paramsArray.SetValue(cancel, 3);
            EventBinding.RaiseCustomEvent("MailMergeBeforeMerge", ref paramsArray);

            cancel = ToBoolean(paramsArray[3]);
        }

        public void MailMergeBeforeRecordMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel)
        {
            if (!Validate("MailMergeBeforeRecordMerge"))
            {
                Invoker.ReleaseParamsArray(doc, cancel);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, NetOffice.WordApi.Document.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newDoc;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("MailMergeBeforeRecordMerge", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        public void MailMergeDataSourceLoad([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("MailMergeDataSourceLoad"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, NetOffice.WordApi.Document.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
            paramsArray[0] = newDoc;
            EventBinding.RaiseCustomEvent("MailMergeDataSourceLoad", ref paramsArray);
        }

        public void MailMergeDataSourceValidate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object handled)
        {
            if (!Validate("MailMergeDataSourceValidate"))
            {
                Invoker.ReleaseParamsArray(doc, handled);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, NetOffice.WordApi.Document.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newDoc;
            paramsArray.SetValue(handled, 1);
            EventBinding.RaiseCustomEvent("MailMergeDataSourceValidate", ref paramsArray);

            handled = ToBoolean(paramsArray[1]);
        }

        public void MailMergeWizardSendToCustom([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("MailMergeWizardSendToCustom"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, NetOffice.WordApi.Document.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
            paramsArray[0] = newDoc;
            EventBinding.RaiseCustomEvent("MailMergeWizardSendToCustom", ref paramsArray);
        }

        public void MailMergeWizardStateChange([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object fromState, [In] [Out] ref object toState, [In] [Out] ref object handled)
        {
            if (!Validate("MailMergeWizardStateChange"))
            {
                Invoker.ReleaseParamsArray(doc, fromState, toState, handled);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, NetOffice.WordApi.Document.LateBindingApiWrapperType);
            object[] paramsArray = new object[4];
            paramsArray[0] = newDoc;
            paramsArray.SetValue(fromState, 1);
            paramsArray.SetValue(toState, 2);
            paramsArray.SetValue(handled, 3);
            EventBinding.RaiseCustomEvent("MailMergeWizardStateChange", ref paramsArray);

            fromState = ToInt32(paramsArray[1]);
            toState = ToInt32(paramsArray[2]);
            handled = ToBoolean(paramsArray[3]);
        }

        public void WindowSize([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("WindowSize"))
            {
                Invoker.ReleaseParamsArray(doc, wn);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, NetOffice.WordApi.Document.LateBindingApiWrapperType);
            NetOffice.WordApi.Window newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Window>(EventClass, wn, NetOffice.WordApi.Window.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newDoc;
            paramsArray[1] = newWn;
            EventBinding.RaiseCustomEvent("WindowSize", ref paramsArray);
        }

        public void XMLSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In, MarshalAs(UnmanagedType.IDispatch)] object oldXMLNode, [In, MarshalAs(UnmanagedType.IDispatch)] object newXMLNode, [In] [Out] ref object reason)
        {
            if (!Validate("XMLSelectionChange"))
            {
                Invoker.ReleaseParamsArray(sel, oldXMLNode, newXMLNode, reason);
                return;
            }

			NetOffice.WordApi.Selection newSel = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Selection>(EventClass, sel, NetOffice.WordApi.Selection.LateBindingApiWrapperType);
			NetOffice.WordApi.XMLNode newOldXMLNode = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.XMLNode>(EventClass, oldXMLNode, NetOffice.WordApi.XMLNode.LateBindingApiWrapperType);
			NetOffice.WordApi.XMLNode newNewXMLNode = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.XMLNode>(EventClass, newXMLNode, NetOffice.WordApi.XMLNode.LateBindingApiWrapperType);
			object[] paramsArray = new object[4];
			paramsArray[0] = newSel;
			paramsArray[1] = newOldXMLNode;
			paramsArray[2] = newNewXMLNode;
			paramsArray.SetValue(reason, 3);
			EventBinding.RaiseCustomEvent("XMLSelectionChange", ref paramsArray);

			reason = ToInt32(paramsArray[3]);
		}

		public void XMLValidationError([In, MarshalAs(UnmanagedType.IDispatch)] object xMLNode)
		{
            if (!Validate("XMLValidationError"))
            {
                Invoker.ReleaseParamsArray(xMLNode);
                return;
            }

			NetOffice.WordApi.XMLNode newXMLNode = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.XMLNode>(EventClass, xMLNode, NetOffice.WordApi.XMLNode.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newXMLNode;
			EventBinding.RaiseCustomEvent("XMLValidationError", ref paramsArray);
		}

		public void DocumentSync([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object syncEventType)
		{
            if (!Validate("DocumentSync"))
            {
                Invoker.ReleaseParamsArray(doc, syncEventType);
                return;
            }

			NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, NetOffice.WordApi.Document.LateBindingApiWrapperType);
			NetOffice.OfficeApi.Enums.MsoSyncEventType newSyncEventType = (NetOffice.OfficeApi.Enums.MsoSyncEventType)syncEventType;
			object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray[1] = newSyncEventType;
			EventBinding.RaiseCustomEvent("DocumentSync", ref paramsArray);
		}
        
        public void EPostageInsertEx([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object cpDeliveryAddrStart, [In] object cpDeliveryAddrEnd, [In] object cpReturnAddrStart, [In] object cpReturnAddrEnd, [In] object xaWidth, [In] object yaHeight, [In] object bstrPrinterName, [In] object bstrPaperFeed, [In] object fPrint, [In] [Out] ref object fCancel)
		{
            if (!Validate("EPostageInsertEx"))
            {
                Invoker.ReleaseParamsArray(doc, cpDeliveryAddrStart, cpDeliveryAddrEnd, cpReturnAddrStart, cpReturnAddrEnd, xaWidth, yaHeight, bstrPrinterName, bstrPaperFeed, fPrint, fCancel);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, NetOffice.WordApi.Document.LateBindingApiWrapperType);
            Int32 newcpDeliveryAddrStart = ToInt32(cpDeliveryAddrStart);
			Int32 newcpDeliveryAddrEnd = ToInt32(cpDeliveryAddrEnd);
			Int32 newcpReturnAddrStart = ToInt32(cpReturnAddrStart);
			Int32 newcpReturnAddrEnd = ToInt32(cpReturnAddrEnd);
			Int32 newxaWidth = ToInt32(xaWidth);
			Int32 newyaHeight = ToInt32(yaHeight);
			string newbstrPrinterName = ToString(bstrPrinterName);
			string newbstrPaperFeed = ToString(bstrPaperFeed);
			bool newfPrint = ToBoolean(fPrint);
			object[] paramsArray = new object[11];
			paramsArray[0] = newDoc;
			paramsArray[1] = newcpDeliveryAddrStart;
			paramsArray[2] = newcpDeliveryAddrEnd;
			paramsArray[3] = newcpReturnAddrStart;
			paramsArray[4] = newcpReturnAddrEnd;
			paramsArray[5] = newxaWidth;
			paramsArray[6] = newyaHeight;
			paramsArray[7] = newbstrPrinterName;
			paramsArray[8] = newbstrPaperFeed;
			paramsArray[9] = newfPrint;
			paramsArray.SetValue(fCancel, 10);
			EventBinding.RaiseCustomEvent("EPostageInsertEx", ref paramsArray);

			fCancel = ToBoolean(paramsArray[10]);
		}

		public void MailMergeDataSourceValidate2([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object handled)
        {
            if (!Validate("MailMergeDataSourceValidate2"))
            {
                Invoker.ReleaseParamsArray(doc, handled);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, NetOffice.WordApi.Document.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray.SetValue(handled, 1);
			EventBinding.RaiseCustomEvent("MailMergeDataSourceValidate2", ref paramsArray);

			handled = ToBoolean(paramsArray[1]);
        }

		public void ProtectedViewWindowOpen([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow)
		{
            if (!Validate("ProtectedViewWindowOpen"))
            {
                Invoker.ReleaseParamsArray(pvWindow);
                return;
            }

			NetOffice.WordApi.ProtectedViewWindow newPvWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.ProtectedViewWindow>(EventClass, pvWindow, NetOffice.WordApi.ProtectedViewWindow.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newPvWindow;
			EventBinding.RaiseCustomEvent("ProtectedViewWindowOpen", ref paramsArray);
		}

		public void ProtectedViewWindowBeforeEdit([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow, [In] [Out] ref object cancel)
        {
            if (!Validate("ProtectedViewWindowBeforeEdit"))
            {
                Invoker.ReleaseParamsArray(pvWindow, cancel);
                return;
            }

            NetOffice.WordApi.ProtectedViewWindow newPvWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.ProtectedViewWindow>(EventClass, pvWindow, NetOffice.WordApi.ProtectedViewWindow.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newPvWindow;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("ProtectedViewWindowBeforeEdit", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

        public void ProtectedViewWindowBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow, [In] object closeReason, [In] [Out] ref object cancel)
		{
            if (!Validate("ProtectedViewWindowBeforeClose"))
            {
                Invoker.ReleaseParamsArray(pvWindow, closeReason, cancel);
                return;
            }

            NetOffice.WordApi.ProtectedViewWindow newPvWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.ProtectedViewWindow>(EventClass, pvWindow, NetOffice.WordApi.ProtectedViewWindow.LateBindingApiWrapperType);
            Int32 newCloseReason = ToInt32(closeReason);
			object[] paramsArray = new object[3];
			paramsArray[0] = newPvWindow;
			paramsArray[1] = newCloseReason;
			paramsArray.SetValue(cancel, 2);
			EventBinding.RaiseCustomEvent("ProtectedViewWindowBeforeClose", ref paramsArray);

			cancel = ToBoolean(paramsArray[2]);
		}

		public void ProtectedViewWindowSize([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow)
        {
            if (!Validate("ProtectedViewWindowSize"))
            {
                Invoker.ReleaseParamsArray(pvWindow);
                return;
            }

            NetOffice.WordApi.ProtectedViewWindow newPvWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.ProtectedViewWindow>(EventClass, pvWindow, NetOffice.WordApi.ProtectedViewWindow.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newPvWindow;
			EventBinding.RaiseCustomEvent("ProtectedViewWindowSize", ref paramsArray);
		}

		public void ProtectedViewWindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow)
		{
            if (!Validate("ProtectedViewWindowActivate"))
            {
                Invoker.ReleaseParamsArray(pvWindow);
                return;
            }

            NetOffice.WordApi.ProtectedViewWindow newPvWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.ProtectedViewWindow>(EventClass, pvWindow, NetOffice.WordApi.ProtectedViewWindow.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newPvWindow;
			EventBinding.RaiseCustomEvent("ProtectedViewWindowActivate", ref paramsArray);
		}

		public void ProtectedViewWindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow)
		{
            if (!Validate("ProtectedViewWindowDeactivate"))
            {
                Invoker.ReleaseParamsArray(pvWindow);
                return;
            }

            NetOffice.WordApi.ProtectedViewWindow newPvWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.ProtectedViewWindow>(EventClass, pvWindow, NetOffice.WordApi.ProtectedViewWindow.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newPvWindow;
			EventBinding.RaiseCustomEvent("ProtectedViewWindowDeactivate", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}