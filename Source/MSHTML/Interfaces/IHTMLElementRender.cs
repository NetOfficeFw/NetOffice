using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IHTMLElementRender 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F669-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLElementRender : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hdc">_RemotableHandle hdc</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 DrawToDC(_RemotableHandle hdc);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrPrinterName">string bstrPrinterName</param>
		/// <param name="hdc">_RemotableHandle hdc</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SetDocumentPrinter(string bstrPrinterName, _RemotableHandle hdc);

		#endregion
	}
}
