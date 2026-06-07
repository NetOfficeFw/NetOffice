// Copyright 2025 Cisco Systems, Inc. All rights reserved.
// Licensed under MIT-style license (see LICENSE.txt file).
//
// Generated code file by Claude Haiku 4.5
//

using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface LabelInfo
	/// SupportByVersion Office, 16
	/// </summary>
	/// <remarks>
	/// Represents the label information data object.
	/// <para>The LabelInfo object can be passed to SetLabel method of SensitivityLabel object.</para>
	/// <para>Docs: <see href="https://learn.microsoft.com/en-us/office/vba/api/office.labelinfo"/></para>
	/// </remarks>
	[SupportByVersion("Office", 16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public class LabelInfo : _IMsoDispObj
	{
		#pragma warning disable

		#region Type Information

		/// <summary>
		/// Instance Type
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
		public override Type InstanceType
		{
			get
			{
				return LateBindingApiWrapperType;
			}
		}

		private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static Type LateBindingApiWrapperType
		{
			get
			{
				if (null == _type)
					_type = typeof(LabelInfo);
				return _type;
			}
		}

		#endregion

		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public LabelInfo(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
		///<param name="comProxy">inner wrapped COM proxy</param>
		public LabelInfo(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
		///<param name="comProxy">inner wrapped COM proxy</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LabelInfo(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
		///<param name="comProxy">inner wrapped COM proxy</param>
		///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LabelInfo(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
		///<param name="comProxy">inner wrapped COM proxy</param>
		///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LabelInfo(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LabelInfo(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LabelInfo() : base()
		{
		}

		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LabelInfo(string progId) : base(progId)
		{
		}

		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Office 16
		/// Get/Set
		/// </summary>
		/// <remarks> Gets or sets the GUID that identifies the action to be performed. </remarks>
		[SupportByVersion("Office", 16)]
		public string ActionId
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ActionId");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ActionId", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 16
		/// Get/Set
		/// </summary>
		/// <remarks> Gets or sets how the label was assigned. </remarks>
		[SupportByVersion("Office", 16)]
		public NetOffice.OfficeApi.Enums.MsoAssignmentMethod AssignmentMethod
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoAssignmentMethod>(this, "AssignmentMethod");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AssignmentMethod", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 16
		/// Get/Set
		/// </summary>
		/// <remarks> Gets or sets content markings value. </remarks>
		[SupportByVersion("Office", 16)]
		public Int32 ContentBits
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ContentBits");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ContentBits", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 16
		/// Get/Set
		/// </summary>
		/// <remarks> Gets or sets whether the label is enabled. </remarks>
		[SupportByVersion("Office", 16)]
		public bool IsEnabled
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsEnabled");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IsEnabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 16
		/// Get/Set
		/// </summary>
		/// <remarks> Gets or sets justification text. Required when downgrading labels. </remarks>
		[SupportByVersion("Office", 16)]
		public string Justification
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Justification");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Justification", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 16
		/// Get/Set
		/// </summary>
		/// <remarks> Gets or sets the GUID of the sensitivity label. </remarks>
		[SupportByVersion("Office", 16)]
		public string LabelId
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "LabelId");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LabelId", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 16
		/// Get/Set
		/// </summary>
		/// <remarks> Gets or sets the display name of the label. </remarks>
		[SupportByVersion("Office", 16)]
		public string LabelName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "LabelName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LabelName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 16
		/// Get/Set
		/// </summary>
		/// <remarks> Gets or sets the date when the label was set. </remarks>
		[SupportByVersion("Office", 16)]
		public string SetDate
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SetDate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SetDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 16
		/// Get/Set
		/// </summary>
		/// <remarks> Gets or sets the GUID of the SharePoint site. </remarks>
		[SupportByVersion("Office", 16)]
		public string SiteId
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SiteId");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SiteId", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
