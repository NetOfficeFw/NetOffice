using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IElementBehaviorLayout 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F6BA-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IElementBehaviorLayout : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="dwFlags">Int32 dwFlags</param>
		/// <param name="sizeContent">tagSIZE sizeContent</param>
		/// <param name="pptTranslateBy">tagPOINT pptTranslateBy</param>
		/// <param name="pptTopLeft">tagPOINT pptTopLeft</param>
		/// <param name="psizeProposed">tagSIZE psizeProposed</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetSize(Int32 dwFlags, tagSIZE sizeContent, tagPOINT pptTranslateBy, tagPOINT pptTopLeft, tagSIZE psizeProposed);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetLayoutInfo();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="lFlags">Int32 lFlags</param>
		/// <param name="pptTopLeft">tagPOINT pptTopLeft</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetPosition(Int32 lFlags, tagPOINT pptTopLeft);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="psizeIn">tagSIZE psizeIn</param>
		/// <param name="prcOut">tagRECT prcOut</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 MapSize(tagSIZE psizeIn, out tagRECT prcOut);

		#endregion
	}
}
