using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Publisher", 14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("00021244-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface DocumentEvents
	{
		[SupportByVersion("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void Open();

		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void BeforeClose([In] [Out] ref object cancel);

		[SupportByVersion("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void ShapesAdded();

		[SupportByVersion("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void WizardAfterChange();

		[SupportByVersion("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
		void ShapesRemoved();

		[SupportByVersion("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void Undo();

		[SupportByVersion("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void Redo();
	}
	
	#endregion
	
	#region SinkHelper
	
    [InternalEntity(InternalEntityKind.SinkHelper)]
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class DocumentEvents_SinkHelper : SinkHelper, DocumentEvents
	{
		#region Static
		
		public static readonly string Id = "00021244-0000-0000-C000-000000000046";
		
		#endregion	
		
		#region Ctor

		public DocumentEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion	

		#region DocumentEvents
		
		public void Open()
		{
            if (!Validate("Open"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Open", ref paramsArray);
		}

		public void BeforeClose([In] [Out] ref object cancel)
		{
            if (!Validate("BeforeClose"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeClose", ref paramsArray);

			cancel = ToBoolean(paramsArray[0]);
		}

		public void ShapesAdded()
		{
            if (!Validate("ShapesAdded"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ShapesAdded", ref paramsArray);
		}

		public void WizardAfterChange()
		{
            if (!Validate("WizardAfterChange"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("WizardAfterChange", ref paramsArray);
		}

		public void ShapesRemoved()
		{
            if (!Validate("ShapesRemoved"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ShapesRemoved", ref paramsArray);
		}

		public void Undo()
		{
            if (!Validate("Undo"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Undo", ref paramsArray);
		}

		public void Redo()
		{
            if (!Validate("Redo"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Redo", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}