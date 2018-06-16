using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface PPDialogs 
	/// SupportByVersion PowerPoint, 9
	/// </summary>
	[SupportByVersion("PowerPoint", 9)]
	[EntityType(EntityType.IsDispatchInterface), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("9149349E-5A91-11CF-8700-00AA0060263B")]
	public interface PPDialogs : Collection
	{
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.Tags Tags { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		string Name { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("PowerPoint", 9)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.PowerPointApi.PPDialog this[object index] { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		/// <param name="position">optional NetOffice.PowerPointApi.Enums.PpDialogPositioning Position = 1</param>
		/// <param name="displayHelp">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayHelp = -1</param>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPDialog AddDialog(Single left, Single top, Single width, Single height, object modal, object parentWindow, object position, object displayHelp);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPDialog AddDialog(Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPDialog AddDialog(Single left, Single top, Single width, Single height, object modal);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPDialog AddDialog(Single left, Single top, Single width, Single height, object modal, object parentWindow);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		/// <param name="position">optional NetOffice.PowerPointApi.Enums.PpDialogPositioning Position = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPDialog AddDialog(Single left, Single top, Single width, Single height, object modal, object parentWindow, object position);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		/// <param name="position">optional NetOffice.PowerPointApi.Enums.PpDialogPositioning Position = 1</param>
		/// <param name="displayHelp">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayHelp = -1</param>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPDialog AddTabDialog(Single left, Single top, Single width, Single height, object modal, object parentWindow, object position, object displayHelp);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPDialog AddTabDialog(Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPDialog AddTabDialog(Single left, Single top, Single width, Single height, object modal);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPDialog AddTabDialog(Single left, Single top, Single width, Single height, object modal, object parentWindow);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		/// <param name="position">optional NetOffice.PowerPointApi.Enums.PpDialogPositioning Position = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPDialog AddTabDialog(Single left, Single top, Single width, Single height, object modal, object parentWindow, object position);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="resourceDLL">string resourceDLL</param>
		/// <param name="nResID">Int32 nResID</param>
		/// <param name="bModal">optional NetOffice.OfficeApi.Enums.MsoTriState bModal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		/// <param name="position">optional NetOffice.PowerPointApi.Enums.PpDialogPositioning Position = 1</param>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPDialog LoadDialog(string resourceDLL, Int32 nResID, object bModal, object parentWindow, object position);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="resourceDLL">string resourceDLL</param>
		/// <param name="nResID">Int32 nResID</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPDialog LoadDialog(string resourceDLL, Int32 nResID);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="resourceDLL">string resourceDLL</param>
		/// <param name="nResID">Int32 nResID</param>
		/// <param name="bModal">optional NetOffice.OfficeApi.Enums.MsoTriState bModal = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPDialog LoadDialog(string resourceDLL, Int32 nResID, object bModal);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="resourceDLL">string resourceDLL</param>
		/// <param name="nResID">Int32 nResID</param>
		/// <param name="bModal">optional NetOffice.OfficeApi.Enums.MsoTriState bModal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPDialog LoadDialog(string resourceDLL, Int32 nResID, object bModal, object parentWindow);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPAlert AddAlert();

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="text">string text</param>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpAlertType type</param>
		/// <param name="icon">NetOffice.PowerPointApi.Enums.PpAlertIcon icon</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.Enums.PpAlertButton RunCharacterAlert(string text, NetOffice.PowerPointApi.Enums.PpAlertType type, NetOffice.PowerPointApi.Enums.PpAlertIcon icon, object parentWindow);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="text">string text</param>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpAlertType type</param>
		/// <param name="icon">NetOffice.PowerPointApi.Enums.PpAlertIcon icon</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.Enums.PpAlertButton RunCharacterAlert(string text, NetOffice.PowerPointApi.Enums.PpAlertType type, NetOffice.PowerPointApi.Enums.PpAlertIcon icon);

		#endregion
	}
}
