using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;
using NetOffice.Exceptions;

namespace NetOffice.PublisherApi.Behind.EventContracts
{
	
	/// <summary>
	/// Default implementation of <see cref="NetOffice.PublisherApi.EventContracts.DocumentEvents"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class DocumentEvents_SinkHelper : SinkHelper, NetOffice.PublisherApi.EventContracts.DocumentEvents
	{
		#region Static
		
		/// <summary>
		/// Interface Id from DocumentEvents
		/// </summary>
		public static readonly string Id = "00021244-0000-0000-C000-000000000046";
		
		#endregion	
		
		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public DocumentEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion	

		#region DocumentEvents
		
		/// <summary>
		/// 
		/// </summary>
		public void Open()
		{
            if (!Validate("Open"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Open", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
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

		/// <summary>
		/// 
		/// </summary>
		public void ShapesAdded()
		{
            if (!Validate("ShapesAdded"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ShapesAdded", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void WizardAfterChange()
		{
            if (!Validate("WizardAfterChange"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("WizardAfterChange", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void ShapesRemoved()
		{
            if (!Validate("ShapesRemoved"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ShapesRemoved", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void Undo()
		{
            if (!Validate("Undo"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Undo", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
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
	
}
