using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLGenericElement 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("3050F4B7-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLGenericElement : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object recordset { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="dataMember">string dataMember</param>
		/// <param name="hierarchy">optional object hierarchy</param>
		[SupportByVersion("MSHTML", 4)]
		object namedRecordset(string dataMember, object hierarchy);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="dataMember">string dataMember</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		object namedRecordset(string dataMember);

		#endregion
	}
}
