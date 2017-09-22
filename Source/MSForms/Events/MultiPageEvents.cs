using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("MSForms", 2)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("7B020EC8-AF6C-11CE-9F46-00AA00574A4F"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface MultiPageEvents
	{
		[SupportByVersion("MSForms", 2)]
        [SinkArgument("index", SinkArgumentType.Int32)]
        [SinkArgument("control", typeof(MSFormsApi.Control))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(768)]
		void AddControl([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object control);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("index", SinkArgumentType.Int32)]
        [SinkArgument("cancel", typeof(MSFormsApi.ReturnBoolean))]
        [SinkArgument("control", typeof(MSFormsApi.Control))]
        [SinkArgument("data", typeof(MSFormsApi.DataObject))]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [SinkArgument("state", SinkArgumentType.Enum, typeof(MSFormsApi.Enums.fmDragState))]
        [SinkArgument("effect", typeof(MSFormsApi.ReturnEffect))]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void BeforeDragOver([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object control, [In, MarshalAs(UnmanagedType.IDispatch)] object data, [In] object x, [In] object y, [In] object state, [In, MarshalAs(UnmanagedType.IDispatch)] object effect, [In] object shift);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("index", SinkArgumentType.Int32)]
        [SinkArgument("cancel", typeof(MSFormsApi.ReturnBoolean))]
        [SinkArgument("control", typeof(MSFormsApi.Control))]
        [SinkArgument("action", SinkArgumentType.Enum, typeof(MSFormsApi.Enums.fmAction))]
        [SinkArgument("data", typeof(MSFormsApi.DataObject))]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [SinkArgument("effect", typeof(MSFormsApi.ReturnEffect))]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void BeforeDropOrPaste([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object control, [In] object action, [In, MarshalAs(UnmanagedType.IDispatch)] object data, [In] object x, [In] object y, [In, MarshalAs(UnmanagedType.IDispatch)] object effect, [In] object shift);

		[SupportByVersion("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void Change();

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("index", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-600)]
		void Click([In] object index);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("index", SinkArgumentType.Int32)]
        [SinkArgument("cancel", typeof(MSFormsApi.ReturnBoolean))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-601)]
		void DblClick([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("index", SinkArgumentType.Int32)]
        [SinkArgument("number", SinkArgumentType.Int32)]
        [SinkArgument("description", typeof(MSFormsApi.ReturnString))]
        [SinkArgument("sCode", SinkArgumentType.Int32)]
        [SinkArgument("source", SinkArgumentType.String)]
        [SinkArgument("helpFile", SinkArgumentType.String)]
        [SinkArgument("helpContext", SinkArgumentType.Int32)]
        [SinkArgument("cancelDisplay", typeof(MSFormsApi.ReturnBoolean))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-608)]
		void Error([In] object index, [In] object number, [In, MarshalAs(UnmanagedType.IDispatch)] object description, [In] object sCode, [In] object source, [In] object helpFile, [In] object helpContext, [In, MarshalAs(UnmanagedType.IDispatch)] object cancelDisplay);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("keyCode", typeof(MSFormsApi.ReturnInteger))]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-602)]
		void KeyDown([In, MarshalAs(UnmanagedType.IDispatch)] object keyCode, [In] object shift);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("keyAscii", typeof(MSFormsApi.ReturnInteger))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-603)]
		void KeyPress([In, MarshalAs(UnmanagedType.IDispatch)] object keyAscii);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("keyCode", typeof(MSFormsApi.ReturnInteger))]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-604)]
		void KeyUp([In, MarshalAs(UnmanagedType.IDispatch)] object keyCode, [In] object shift);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("index", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(770)]
		void Layout([In] object index);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("index", SinkArgumentType.Int32)]
        [SinkArgument("button", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-605)]
		void MouseDown([In] object index, [In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("index", SinkArgumentType.Int32)]
        [SinkArgument("button", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-606)]
		void MouseMove([In] object index, [In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("index", SinkArgumentType.Int32)]
        [SinkArgument("button", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-607)]
		void MouseUp([In] object index, [In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("index", SinkArgumentType.Int32)]
        [SinkArgument("control", typeof(MSFormsApi.Control))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(771)]
		void RemoveControl([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object control);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("index", SinkArgumentType.Int32)]
        [SinkArgument("actionX", SinkArgumentType.Enum, typeof(MSFormsApi.Enums.fmScrollAction))]
        [SinkArgument("actionY", SinkArgumentType.Enum, typeof(MSFormsApi.Enums.fmScrollAction))]
        [SinkArgument("requestDx", SinkArgumentType.Single)]
        [SinkArgument("requestDy", SinkArgumentType.Single)]
        [SinkArgument("actualDx", typeof(MSFormsApi.ReturnSingle))]
        [SinkArgument("actualDy", typeof(MSFormsApi.ReturnSingle))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(772)]
		void Scroll([In] object index, [In] object actionX, [In] object actionY, [In] object requestDx, [In] object requestDy, [In, MarshalAs(UnmanagedType.IDispatch)] object actualDx, [In, MarshalAs(UnmanagedType.IDispatch)] object actualDy);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("index", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(773)]
		void Zoom([In] object index, [In] [Out] ref object percent);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class MultiPageEvents_SinkHelper : SinkHelper, MultiPageEvents
	{
		#region Static
		
		public static readonly string Id = "7B020EC8-AF6C-11CE-9F46-00AA00574A4F";
		
		#endregion
	
		#region Ctor

		public MultiPageEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region MultiPageEvents
		
        public void AddControl([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object control)
		{
            if (!Validate("AddControl"))
            {
                Invoker.ReleaseParamsArray(index, control);
                return;
            }

			Int32 newIndex = ToInt32(index);
			NetOffice.MSFormsApi.Control newControl = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.Control>(EventClass, control, NetOffice.MSFormsApi.Control.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newIndex;
			paramsArray[1] = newControl;
			EventBinding.RaiseCustomEvent("AddControl", ref paramsArray);
		}

        public void BeforeDragOver([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object control, [In, MarshalAs(UnmanagedType.IDispatch)] object data, [In] object x, [In] object y, [In] object state, [In, MarshalAs(UnmanagedType.IDispatch)] object effect, [In] object shift)
        {
            if (!Validate("BeforeDragOver"))
            {
                Invoker.ReleaseParamsArray(index, cancel, control, data, x, y, state, effect, shift);
                return;
            }

			Int32 newIndex = Convert.ToInt32(index);
			NetOffice.MSFormsApi.ReturnBoolean newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnBoolean>(EventClass, cancel, NetOffice.MSFormsApi.ReturnBoolean.LateBindingApiWrapperType);
			NetOffice.MSFormsApi.Control newControl = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.Control>(EventClass, control, NetOffice.MSFormsApi.Control.LateBindingApiWrapperType);
			NetOffice.MSFormsApi.DataObject newData = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.DataObject>(EventClass, data, NetOffice.MSFormsApi.DataObject.LateBindingApiWrapperType);
			Single newX = ToSingle(x);
			Single newY = ToSingle(y);
			NetOffice.MSFormsApi.Enums.fmDragState newState = (NetOffice.MSFormsApi.Enums.fmDragState)state;
			NetOffice.MSFormsApi.ReturnEffect newEffect = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnEffect>(EventClass, effect, NetOffice.MSFormsApi.ReturnEffect.LateBindingApiWrapperType);
			Int16 newShift = ToInt16(shift);
			object[] paramsArray = new object[9];
			paramsArray[0] = newIndex;
			paramsArray[1] = newCancel;
			paramsArray[2] = newControl;
			paramsArray[3] = newData;
			paramsArray[4] = newX;
			paramsArray[5] = newY;
			paramsArray[6] = newState;
			paramsArray[7] = newEffect;
			paramsArray[8] = newShift;
			EventBinding.RaiseCustomEvent("BeforeDragOver", ref paramsArray);
		}

        public void BeforeDropOrPaste([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object control, [In] object action, [In, MarshalAs(UnmanagedType.IDispatch)] object data, [In] object x, [In] object y, [In, MarshalAs(UnmanagedType.IDispatch)] object effect, [In] object shift)
        {
            if (!Validate("BeforeDropOrPaste"))
            {
                Invoker.ReleaseParamsArray(index, cancel, control, action, data, x, y, effect, shift);
                return;
            }

			Int32 newIndex = ToInt32(index);
			NetOffice.MSFormsApi.ReturnBoolean newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnBoolean>(EventClass, cancel, NetOffice.MSFormsApi.ReturnBoolean.LateBindingApiWrapperType);
			NetOffice.MSFormsApi.Control newControl = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.Control>(EventClass, control, NetOffice.MSFormsApi.Control.LateBindingApiWrapperType);
			NetOffice.MSFormsApi.Enums.fmAction newAction = (NetOffice.MSFormsApi.Enums.fmAction)action;
			NetOffice.MSFormsApi.DataObject newData = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.DataObject>(EventClass, data, NetOffice.MSFormsApi.DataObject.LateBindingApiWrapperType);
			Single newX = ToSingle(x);
			Single newY = ToSingle(y);
			NetOffice.MSFormsApi.ReturnEffect newEffect = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnEffect>(EventClass, effect, NetOffice.MSFormsApi.ReturnEffect.LateBindingApiWrapperType);
			Int16 newShift = ToInt16(shift);
			object[] paramsArray = new object[9];
			paramsArray[0] = newIndex;
			paramsArray[1] = newCancel;
			paramsArray[2] = newControl;
			paramsArray[3] = newAction;
			paramsArray[4] = newData;
			paramsArray[5] = newX;
			paramsArray[6] = newY;
			paramsArray[7] = newEffect;
			paramsArray[8] = newShift;
			EventBinding.RaiseCustomEvent("BeforeDropOrPaste", ref paramsArray);
		}

		public void Change()
        {
            if (!Validate("Change"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Change", ref paramsArray);
		}

        public void Click([In] object index)
		{
            if (!Validate("Click"))
            {
                return;
            }

			Int32 newIndex = ToInt32(index);
			object[] paramsArray = new object[1];
			paramsArray[0] = newIndex;
			EventBinding.RaiseCustomEvent("Click", ref paramsArray);
		}

        public void DblClick([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
		{
            if (!Validate("DblClick"))
            {
                Invoker.ReleaseParamsArray(index, cancel);
                return;
            }

			Int32 newIndex = Convert.ToInt32(index);
			NetOffice.MSFormsApi.ReturnBoolean newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnBoolean>(EventClass, cancel, NetOffice.MSFormsApi.ReturnBoolean.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newIndex;
			paramsArray[1] = newCancel;
			EventBinding.RaiseCustomEvent("DblClick", ref paramsArray);
		}

        public void Error([In] object index, [In] object number, [In, MarshalAs(UnmanagedType.IDispatch)] object description, [In] object sCode, [In] object source, [In] object helpFile, [In] object helpContext, [In, MarshalAs(UnmanagedType.IDispatch)] object cancelDisplay)
        {
            if (!Validate("Error"))
            {
                Invoker.ReleaseParamsArray(index, number, description, sCode, source, helpFile, helpContext, cancelDisplay);
                return;
            }

			Int32 newIndex = ToInt32(index);
			Int16 newNumber = ToInt16(number);
			NetOffice.MSFormsApi.ReturnString newDescription = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnString>(EventClass, description, NetOffice.MSFormsApi.ReturnString.LateBindingApiWrapperType);
			Int32 newSCode = ToInt32(sCode);
			string newSource = ToString(source);
			string newHelpFile = ToString(helpFile);
			Int32 newHelpContext = ToInt32(helpContext);
			NetOffice.MSFormsApi.ReturnBoolean newCancelDisplay = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnBoolean>(EventClass, cancelDisplay, NetOffice.MSFormsApi.ReturnBoolean.LateBindingApiWrapperType);
			object[] paramsArray = new object[8];
			paramsArray[0] = newIndex;
			paramsArray[1] = newNumber;
			paramsArray[2] = newDescription;
			paramsArray[3] = newSCode;
			paramsArray[4] = newSource;
			paramsArray[5] = newHelpFile;
			paramsArray[6] = newHelpContext;
			paramsArray[7] = newCancelDisplay;
			EventBinding.RaiseCustomEvent("Error", ref paramsArray);
		}

        public void KeyDown([In, MarshalAs(UnmanagedType.IDispatch)] object keyCode, [In] object shift)
        {
            if (!Validate("KeyDown"))
            {
                Invoker.ReleaseParamsArray(keyCode, shift);
                return;
            }

			NetOffice.MSFormsApi.ReturnInteger newKeyCode = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnInteger>(EventClass, keyCode, NetOffice.MSFormsApi.ReturnInteger.LateBindingApiWrapperType);
			Int16 newShift = ToInt16(shift);
			object[] paramsArray = new object[2];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newShift;
			EventBinding.RaiseCustomEvent("KeyDown", ref paramsArray);
		}

        public void KeyPress([In, MarshalAs(UnmanagedType.IDispatch)] object keyAscii)
        {
            if (!Validate("KeyPress"))
            {
                Invoker.ReleaseParamsArray(keyAscii);
                return;
            }

			NetOffice.MSFormsApi.ReturnInteger newKeyAscii = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnInteger>(EventClass, keyAscii, NetOffice.MSFormsApi.ReturnInteger.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newKeyAscii;
			EventBinding.RaiseCustomEvent("KeyPress", ref paramsArray);
		}

        public void KeyUp([In, MarshalAs(UnmanagedType.IDispatch)] object keyCode, [In] object shift)
        {
            if (!Validate("KeyUp"))
            {
                Invoker.ReleaseParamsArray(keyCode, shift);
                return;
            }

			NetOffice.MSFormsApi.ReturnInteger newKeyCode = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnInteger>(EventClass, keyCode, NetOffice.MSFormsApi.ReturnInteger.LateBindingApiWrapperType);
			Int16 newShift = Convert.ToInt16(shift);
			object[] paramsArray = new object[2];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newShift;
			EventBinding.RaiseCustomEvent("KeyUp", ref paramsArray);
		}

        public void Layout([In] object index)
        {
            if (!Validate("Layout"))
            {
                Invoker.ReleaseParamsArray(index);
                return;
            }

			Int32 newIndex = ToInt32(index);
			object[] paramsArray = new object[1];
			paramsArray[0] = newIndex;
			EventBinding.RaiseCustomEvent("Layout", ref paramsArray);
		}

        public void MouseDown([In] object index, [In] object button, [In] object shift, [In] object x, [In] object y)
        {
            if (!Validate("MouseDown"))
            {
                Invoker.ReleaseParamsArray(index, button, shift, x, y);
                return;
            }

			Int32 newIndex = ToInt32(index);
			Int16 newButton = ToInt16(button);
			Int16 newShift = ToInt16(shift);
			Single newX = ToSingle(x);
			Single newY = ToSingle(y);
			object[] paramsArray = new object[5];
			paramsArray[0] = newIndex;
			paramsArray[1] = newButton;
			paramsArray[2] = newShift;
			paramsArray[3] = newX;
			paramsArray[4] = newY;
			EventBinding.RaiseCustomEvent("MouseDown", ref paramsArray);
		}

        public void MouseMove([In] object index, [In] object button, [In] object shift, [In] object x, [In] object y)
		{
            if (!Validate("MouseMove"))
            {
                Invoker.ReleaseParamsArray(index, button, shift, x, y);
                return;
            }

            Int32 newIndex = ToInt32(index);
			Int16 newButton = ToInt16(button);
			Int16 newShift = ToInt16(shift);
			Single newX = ToSingle(x);
			Single newY = ToSingle(y);
			object[] paramsArray = new object[5];
			paramsArray[0] = newIndex;
			paramsArray[1] = newButton;
			paramsArray[2] = newShift;
			paramsArray[3] = newX;
			paramsArray[4] = newY;
			EventBinding.RaiseCustomEvent("MouseMove", ref paramsArray);
		}

        public void MouseUp([In] object index, [In] object button, [In] object shift, [In] object x, [In] object y)
		{
            if (!Validate("MouseUp"))
            {
                Invoker.ReleaseParamsArray(index, button, shift, x, y);
                return;
            }

            Int32 newIndex = ToInt32(index);
			Int16 newButton = ToInt16(button);
			Int16 newShift = ToInt16(shift);
			Single newX = ToSingle(x);
			Single newY = ToSingle(y);
			object[] paramsArray = new object[5];
			paramsArray[0] = newIndex;
			paramsArray[1] = newButton;
			paramsArray[2] = newShift;
			paramsArray[3] = newX;
			paramsArray[4] = newY;
			EventBinding.RaiseCustomEvent("MouseUp", ref paramsArray);
		}

        public void RemoveControl([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object control)
        {
            if (!Validate("RemoveControl"))
            {
                Invoker.ReleaseParamsArray(index, control);
                return;
            }

			Int32 newIndex = ToInt32(index);
			NetOffice.MSFormsApi.Control newControl = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.Control>(EventClass, control, NetOffice.MSFormsApi.Control.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newIndex;
			paramsArray[1] = newControl;
			EventBinding.RaiseCustomEvent("RemoveControl", ref paramsArray);
		}

        public void Scroll([In] object index, [In] object actionX, [In] object actionY, [In] object requestDx, [In] object requestDy, [In, MarshalAs(UnmanagedType.IDispatch)] object actualDx, [In, MarshalAs(UnmanagedType.IDispatch)] object actualDy)
        {
            if (!Validate("Scroll"))
            {
                Invoker.ReleaseParamsArray(index, actionX, actionY, requestDx, requestDy, actualDx, actualDy);
                return;
            }

			Int32 newIndex = Convert.ToInt32(index);
			NetOffice.MSFormsApi.Enums.fmScrollAction newActionX = (NetOffice.MSFormsApi.Enums.fmScrollAction)actionX;
			NetOffice.MSFormsApi.Enums.fmScrollAction newActionY = (NetOffice.MSFormsApi.Enums.fmScrollAction)actionY;
			Single newRequestDx = ToSingle(requestDx);
			Single newRequestDy = ToSingle(requestDy);
			NetOffice.MSFormsApi.ReturnSingle newActualDx = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnSingle>(EventClass, actualDx, NetOffice.MSFormsApi.ReturnSingle.LateBindingApiWrapperType);
			NetOffice.MSFormsApi.ReturnSingle newActualDy = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnSingle>(EventClass, actualDy, NetOffice.MSFormsApi.ReturnSingle.LateBindingApiWrapperType);
			object[] paramsArray = new object[7];
			paramsArray[0] = newIndex;
			paramsArray[1] = newActionX;
			paramsArray[2] = newActionY;
			paramsArray[3] = newRequestDx;
			paramsArray[4] = newRequestDy;
			paramsArray[5] = newActualDx;
			paramsArray[6] = newActualDy;
			EventBinding.RaiseCustomEvent("Scroll", ref paramsArray);
		}

        public void Zoom([In] object index, [In] [Out] ref object percent)
		{
            if (!Validate("Zoom"))
            {
                Invoker.ReleaseParamsArray(index, percent);
                return;
            }

			Int32 newIndex = ToInt32(index);
			object[] paramsArray = new object[2];
			paramsArray[0] = newIndex;
			paramsArray.SetValue(percent, 1);
			EventBinding.RaiseCustomEvent("Zoom", ref paramsArray);

            percent = ToInt32(paramsArray[1]);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}