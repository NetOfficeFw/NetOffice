using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface ProtectedViewWindow 
	/// SupportByVersion Word, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838924.aspx </remarks>
	[SupportByVersion("Word", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("F743EDD0-9B97-4B09-89CC-77BE19B51481")]
	public interface ProtectedViewWindow : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823237.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196234.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196844.aspx </remarks>
		[SupportByVersion("Word", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840025.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		string Caption { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835737.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.Document Document { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196080.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Int32 Left { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193919.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Int32 Top { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196954.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Int32 Width { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837884.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Int32 Height { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822310.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.Enums.WdWindowState WindowState { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822915.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		bool Active { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838301.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Int32 Index { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194685.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		bool Visible { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820906.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		string SourceName { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845066.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		string SourcePath { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835830.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		void Activate();

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822887.aspx </remarks>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.Document Edit(object passwordTemplate, object writePasswordDocument, object writePasswordTemplate);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822887.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.Document Edit();

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822887.aspx </remarks>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.Document Edit(object passwordTemplate);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822887.aspx </remarks>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.Document Edit(object passwordTemplate, object writePasswordDocument);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194469.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		void Close();

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821091.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		void ToggleRibbon();

		#endregion
	}
}
