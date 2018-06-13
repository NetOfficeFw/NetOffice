using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface Broadcast 
	/// SupportByVersion Word, 15, 16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229208.aspx </remarks>
	[SupportByVersion("Word", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("B67DE22C-BC01-4A73-A99B-070D1B5A795D")]
	public interface Broadcast : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231615.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231440.aspx </remarks>
		[SupportByVersion("Word", 15, 16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228838.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		string AttendeeUrl { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231748.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		NetOffice.OfficeApi.Enums.MsoBroadcastState State { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230812.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		Int32 Capabilities { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228473.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		string PresenterServiceUrl { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230720.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		string SessionID { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227544.aspx </remarks>
		/// <param name="serverUrl">string serverUrl</param>
		[SupportByVersion("Word", 15, 16)]
		void Start(string serverUrl);

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228106.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		void Pause();

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230520.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		void Resume();

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232060.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		void End();

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232140.aspx </remarks>
		/// <param name="notesUrl">string notesUrl</param>
		/// <param name="notesWacUrl">string notesWacUrl</param>
		[SupportByVersion("Word", 15, 16)]
		void AddMeetingNotes(string notesUrl, string notesWacUrl);

		#endregion
	}
}
