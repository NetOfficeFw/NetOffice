using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLOptionsHolder 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("3050F378-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLOptionsHolder : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.MSHTMLApi.IHTMLDocument2 document { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.MSHTMLApi.IHTMLFontNamesCollection fonts { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		object execArg { get; set; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 errorLine { get; set; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 errorCharacter { get; set; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 errorCode { get; set; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		string errorMessage { get; set; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		bool errorDebug { get; set; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.MSHTMLApi.IHTMLWindow2 unsecuredWindowOfDocument { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		string findText { get; set; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		bool anythingAfterFrameset { get; set; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		string secureConnectionInfo { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fontName">string fontName</param>
		[SupportByVersion("MSHTML", 4)]
		NetOffice.MSHTMLApi.IHTMLFontSizesCollection sizes(string fontName);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		/// <param name="initDir">optional object initDir</param>
		/// <param name="filter">optional object filter</param>
		/// <param name="title">optional object title</param>
		[SupportByVersion("MSHTML", 4)]
		string openfiledlg(object initFile, object initDir, object filter, object title);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		string openfiledlg();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		string openfiledlg(object initFile);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		/// <param name="initDir">optional object initDir</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		string openfiledlg(object initFile, object initDir);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		/// <param name="initDir">optional object initDir</param>
		/// <param name="filter">optional object filter</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		string openfiledlg(object initFile, object initDir, object filter);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		/// <param name="initDir">optional object initDir</param>
		/// <param name="filter">optional object filter</param>
		/// <param name="title">optional object title</param>
		[SupportByVersion("MSHTML", 4)]
		string savefiledlg(object initFile, object initDir, object filter, object title);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		string savefiledlg();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		string savefiledlg(object initFile);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		/// <param name="initDir">optional object initDir</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		string savefiledlg(object initFile, object initDir);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initFile">optional object initFile</param>
		/// <param name="initDir">optional object initDir</param>
		/// <param name="filter">optional object filter</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		string savefiledlg(object initFile, object initDir, object filter);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initColor">optional object initColor</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 choosecolordlg(object initColor);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		Int32 choosecolordlg();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		void showSecurityInfo();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="_object">NetOffice.MSHTMLApi.IHTMLObjectElement object</param>
		[SupportByVersion("MSHTML", 4)]
		bool isApartmentModel(NetOffice.MSHTMLApi.IHTMLObjectElement _object);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fontName">string fontName</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 getCharset(string fontName);

		#endregion
	}
}
