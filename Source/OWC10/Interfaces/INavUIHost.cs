using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// Interface INavUIHost 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("F5B39AC5-1480-11D3-8549-00C04FAC67D7")]
	public interface INavUIHost : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="navbtn">Int32 navbtn</param>
		[SupportByVersion("OWC10", 1)]
		Int32 IsButtonEnabled(Int32 navbtn);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="navbtn">Int32 navbtn</param>
		/// <param name="cancel">Int32 cancel</param>
		[SupportByVersion("OWC10", 1)]
		Int32 BeforeButtonClick(Int32 navbtn, Int32 cancel);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="navbtn">Int32 navbtn</param>
		[SupportByVersion("OWC10", 1)]
		Int32 AfterButtonClick(Int32 navbtn);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="displayText">string displayText</param>
		[SupportByVersion("OWC10", 1)]
		Int32 GetDisplayText(string displayText);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 OnNavUIChange();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 IsFilterOn();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 IsContextBiDi();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="fontName">string fontName</param>
		[SupportByVersion("OWC10", 1)]
		Int32 GetFontName(string fontName);

		#endregion
	}
}
