using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.MSComctlLibApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("MSComctlLib", 6)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("2C247F22-8591-11D1-B16A-00C0F0283628"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ImageListEvents
	{
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ImageListEvents_SinkHelper : SinkHelper, ImageListEvents
	{
		#region Static
		
		public static readonly string Id = "2C247F22-8591-11D1-B16A-00C0F0283628";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private ICOMObject _eventClass;
        
		#endregion
		
		#region Ctor

		public ImageListEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region ImageListEvents
		
		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}