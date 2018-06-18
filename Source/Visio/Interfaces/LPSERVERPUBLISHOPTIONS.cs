using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// Interface LPSERVERPUBLISHOPTIONS 
	/// SupportByVersion Visio, 14,15,16
	/// </summary>
	[SupportByVersion("Visio", 14,15,16)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("00000000-0000-0000-0000-000000000000")]
	public interface LPSERVERPUBLISHOPTIONS : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVApplication Application { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		Int16 Stat { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVDocument Document { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		Int16 ObjectType { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="pageName">string pageName</param>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisLangFlags flags</param>
		[SupportByVersion("Visio", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool get_IsPublishedPage(string pageName, NetOffice.VisioApi.Enums.VisLangFlags flags);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Alias for get_IsPublishedPage
		/// </summary>
		/// <param name="pageName">string pageName</param>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisLangFlags flags</param>
		[SupportByVersion("Visio", 14,15,16), Redirect("get_IsPublishedPage")]
		bool IsPublishedPage(string pageName, NetOffice.VisioApi.Enums.VisLangFlags flags);

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="pageName">string pageName</param>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisLangFlags flags</param>
		[SupportByVersion("Visio", 14,15,16)]
		void IncludePage(string pageName, NetOffice.VisioApi.Enums.VisLangFlags flags);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="pageName">string pageName</param>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisLangFlags flags</param>
		[SupportByVersion("Visio", 14,15,16)]
		void ExcludePage(string pageName, NetOffice.VisioApi.Enums.VisLangFlags flags);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="publishPages">NetOffice.VisioApi.Enums.VisPublishPages publishPages</param>
		/// <param name="namesArray">String[] namesArray</param>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisLangFlags flags</param>
		[SupportByVersion("Visio", 14,15,16)]
		void SetPagesToPublish(NetOffice.VisioApi.Enums.VisPublishPages publishPages, String[] namesArray, NetOffice.VisioApi.Enums.VisLangFlags flags);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisLangFlags flags</param>
		/// <param name="publishPages">NetOffice.VisioApi.Enums.VisPublishPages publishPages</param>
		/// <param name="namesArray">String[] namesArray</param>
		[SupportByVersion("Visio", 14,15,16)]
		void GetPagesToPublish(NetOffice.VisioApi.Enums.VisLangFlags flags, out NetOffice.VisioApi.Enums.VisPublishPages publishPages, out String[] namesArray);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="publishDataRecordsets">NetOffice.VisioApi.Enums.VisPublishDataRecordsets publishDataRecordsets</param>
		/// <param name="dataRecordsetIDs">Int32[] dataRecordsetIDs</param>
		[SupportByVersion("Visio", 14,15,16)]
		void SetRecordsetsToPublish(NetOffice.VisioApi.Enums.VisPublishDataRecordsets publishDataRecordsets, Int32[] dataRecordsetIDs);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="publishDataRecordsets">NetOffice.VisioApi.Enums.VisPublishDataRecordsets publishDataRecordsets</param>
		/// <param name="dataRecordsetIDs">Int32[] dataRecordsetIDs</param>
		[SupportByVersion("Visio", 14,15,16)]
		void GetRecordsetsToPublish(out NetOffice.VisioApi.Enums.VisPublishDataRecordsets publishDataRecordsets, out Int32[] dataRecordsetIDs);

		#endregion
	}
}
