using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi.EventContracts
{
    /// <summary>
    /// DocumentEvents
    /// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("00021244-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface DocumentEvents
	{
		/// <summary>
		/// Open
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void Open();

		/// <summary>
		/// BeforeClose
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void BeforeClose([In] [Out] ref object cancel);

		/// <summary>
		/// ShapesAdded
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void ShapesAdded();

		/// <summary>
		/// WizardAfterChange
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void WizardAfterChange();

		/// <summary>
		/// ShapesRemoved
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
		void ShapesRemoved();

		/// <summary>
		/// Undo
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void Undo();

		/// <summary>
		/// Redo
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void Redo();
	}	
}
