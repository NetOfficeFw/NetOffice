using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _Inspector 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("00063005-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi.Inspector))]
    public interface _Inspector : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868774.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Application Application { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868246.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlObjectClass Class { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869600.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._NameSpace Session { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866024.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.CommandBars CommandBars { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869722.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		object CurrentItem { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868598.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlEditorType EditorType { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868497.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		object ModifiedFormPages { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		object HTMLEditor { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868196.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		object WordEditor { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870165.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Caption { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867136.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		Int32 Height { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865095.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		Int32 Left { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866432.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		Int32 Top { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867846.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		Int32 Width { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869599.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlWindowState WindowState { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861823.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._AttachmentSelection AttachmentSelection { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869503.aspx </remarks>
		/// <param name="saveMode">NetOffice.OutlookApi.Enums.OlInspectorClose saveMode</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void Close(NetOffice.OutlookApi.Enums.OlInspectorClose saveMode);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864462.aspx </remarks>
		/// <param name="modal">optional object modal</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void Display(object modal);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864462.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void Display();

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870121.aspx </remarks>
		/// <param name="pageName">string pageName</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void HideFormPage(string pageName);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870014.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		bool IsWordMail();

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868251.aspx </remarks>
		/// <param name="pageName">string pageName</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void SetCurrentFormPage(string pageName);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869276.aspx </remarks>
		/// <param name="pageName">string pageName</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void ShowFormPage(string pageName);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869361.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void Activate();

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867657.aspx </remarks>
		/// <param name="control">object control</param>
		/// <param name="propertyName">string propertyName</param>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		void SetControlItemProperty(object control, string propertyName);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868411.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		object NewFormRegion();

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868989.aspx </remarks>
		/// <param name="path">string path</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		object OpenFormRegion(string path);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867614.aspx </remarks>
		/// <param name="page">object page</param>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void SaveFormRegion(object page, string fileName);

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862390.aspx </remarks>
		/// <param name="start">DateTime start</param>
		[SupportByVersion("Outlook", 14,15,16)]
		void SetSchedulingStartTime(DateTime start);

		#endregion
	}
}
