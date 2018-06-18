using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;
using NetOffice.Exceptions;

namespace NetOffice.MSHTMLApi.Behind.EventContracts
{

	/// <summary>
	/// Default implementation of <see cref="NetOffice.MSHTMLApi.EventContracts.HTMLObjectElementEvents"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class HTMLObjectElementEvents_SinkHelper : SinkHelper, NetOffice.MSHTMLApi.EventContracts.HTMLObjectElementEvents
	{
		#region Static
		
		/// <summary>
		/// Interface Id from HTMLObjectElementEvents
		/// </summary>
		public static readonly string Id = "3050F3C4-98B5-11CF-BB82-00AA00BDCE0B";
		
		#endregion
	
		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public HTMLObjectElementEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region HTMLObjectElementEvents
		
		/// <summary>
		/// 
		/// </summary>
		public void onbeforeupdate()
		{
            if (!Validate("onbeforeupdate"))
            {
                return;
            }

			Delegate[] recipients = EventBinding.GetEventRecipients("onbeforeupdate");
			if( (true == EventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onbeforeupdate", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onafterupdate()
		{
            if (!Validate("onafterupdate"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onafterupdate", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onerrorupdate()
		{
            if (!Validate("onerrorupdate"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onerrorupdate", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onrowexit()
		{
            if (!Validate("onrowexit"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onrowexit", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onrowenter()
		{
            if (!Validate("onrowenter"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onrowenter", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void ondatasetchanged()
		{
            if (!Validate("ondatasetchanged"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ondatasetchanged", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void ondataavailable()
		{
            if (!Validate("ondataavailable"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ondataavailable", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void ondatasetcomplete()
		{
            if (!Validate("ondatasetcomplete"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ondatasetcomplete", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onerror()
		{
            if (!Validate("onerror"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onerror", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onrowsdelete()
		{
            if (!Validate("onrowsdelete"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onrowsdelete", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onrowsinserted()
		{
            if (!Validate("onrowsinserted"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onrowsinserted", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void oncellchange()
		{
            if (!Validate("oncellchange"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("oncellchange", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onreadystatechange()
		{
            if (!Validate("onreadystatechange"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onreadystatechange", ref paramsArray);
		}

		#endregion
	}
	
}
