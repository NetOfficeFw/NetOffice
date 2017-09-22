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

	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000209FE-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ApplicationEvents2
	{
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void Startup();

		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void Quit();

		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void DocumentChange();

		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void DocumentOpen([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void DocumentBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void DocumentBeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("saveAsUI", SinkArgumentType.Bool)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void DocumentBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object saveAsUI, [In] [Out] ref object cancel);

		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
		void NewDocument([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("wn", typeof(WordApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10)]
		void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("wn", typeof(WordApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(11)]
		void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
        [SinkArgument("sel", typeof(WordApi.Selection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12)]
		void WindowSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sel);

		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
        [SinkArgument("sel", typeof(WordApi.Selection))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
		void WindowBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel);

		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
        [SinkArgument("sel", typeof(WordApi.Selection))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(14)]
		void WindowBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel);
	}
	
	#endregion
	
	#region SinkHelper

	[InternalEntity(InternalEntityKind.SinkHelper)]
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ApplicationEvents2_SinkHelper : SinkHelper, ApplicationEvents2
	{
		#region Static
		
		public static readonly string Id = "000209FE-0000-0000-C000-000000000046";
		
		#endregion

		#region Ctor

		public ApplicationEvents2_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region ApplicationEvents2
		
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

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}