using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IElementBehaviorLayout2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F846-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IElementBehaviorLayout2 : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="plDescent">Int32 plDescent</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetTextDescent(out Int32 plDescent);

		#endregion
	}
}
