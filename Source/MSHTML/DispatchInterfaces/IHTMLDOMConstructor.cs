using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLDOMConstructor 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("3051049B-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLDOMConstructor : IHTMLStyle6
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="propname">string propname</param>
		[SupportByVersion("MSHTML", 4)]
		object LookupGetter(string propname);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="propname">string propname</param>
		[SupportByVersion("MSHTML", 4)]
		object LookupSetter(string propname);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="propname">string propname</param>
		/// <param name="pdispHandler">object pdispHandler</param>
		[SupportByVersion("MSHTML", 4)]
		void DefineGetter(string propname, object pdispHandler);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="propname">string propname</param>
		/// <param name="pdispHandler">object pdispHandler</param>
		[SupportByVersion("MSHTML", 4)]
		void DefineSetter(string propname, object pdispHandler);

		#endregion
	}
}
