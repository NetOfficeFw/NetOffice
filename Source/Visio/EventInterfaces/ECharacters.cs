using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.VisioApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
	[ComImport, Guid("000D0B0C-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ECharacters
	{
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8320)]
		void TextChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ECharacters_SinkHelper : SinkHelper, ECharacters
	{
		#region Static
		
		public static readonly string Id = "000D0B0C-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public ECharacters_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region Properties

        internal Core Factory
        {
            get
            {
                if (null != _eventClass)
                    return _eventClass.Factory;
                else
                    return Core.Default;
            }
        }

        #endregion

		#region ECharacters Members
		
		public void TextChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("TextChanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(shape);
				return;
			}

			NetOffice.VisioApi.IVShape newShape = Factory.CreateObjectFromComProxy(_eventClass, shape) as NetOffice.VisioApi.IVShape;
			object[] paramsArray = new object[1];
			paramsArray[0] = newShape;
			_eventBinding.RaiseCustomEvent("TextChanged", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}