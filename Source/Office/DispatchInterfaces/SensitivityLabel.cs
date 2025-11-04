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
	/// DispatchInterface SensitivityLabel
	/// SupportByVersion Office, 16
	/// </summary>
	/// <remarks>
	/// Represents a wrapper object for accessing sensitivity label on the active document.
	/// <para>SensitivityLabel applied on a document requires the use of policy defined by organization administrator. The organization is identified by using an identity of an Office Account signed into Office.</para>
	/// <para>Docs: <see href="https://learn.microsoft.com/en-us/office/vba/api/office.sensitivitylabel"/></para>
	/// </remarks>
	[SupportByVersion("Office", 16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public class SensitivityLabel : _IMsoDispObj
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
					_type = typeof(SensitivityLabel);
				return _type;
			}
		}

		#endregion

		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public SensitivityLabel(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
		///<param name="comProxy">inner wrapped COM proxy</param>
		public SensitivityLabel(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
		///<param name="comProxy">inner wrapped COM proxy</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SensitivityLabel(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
		///<param name="comProxy">inner wrapped COM proxy</param>
		///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SensitivityLabel(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
		///<param name="comProxy">inner wrapped COM proxy</param>
		///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SensitivityLabel(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SensitivityLabel(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SensitivityLabel() : base()
		{
		}

		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SensitivityLabel(string progId) : base(progId)
		{
		}

		#endregion

		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 16
		/// </summary>
		/// <remarks>
		/// Gets the current label information that exists on the document for the user.
		/// <para>If the SensitivityLabelPolicy.CompleteInitialize was called, it gets the label for the user that was passed with UserId, otherwise gets the label for the user which is authenticated to the document.</para>
		/// <para>Docs: <see href="https://learn.microsoft.com/en-us/office/vba/api/office.sensitivitylabel.getlabel"/> </remarks>
		[SupportByVersion("Office", 16)]
		public NetOffice.OfficeApi.LabelInfo GetLabel()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.LabelInfo>(this, "GetLabel", NetOffice.OfficeApi.LabelInfo.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Office 16
		/// </summary>
		/// <remarks>
		/// Sets the label information on the document for the user.
		/// <para>If the SensitivityLabelPolicy.CompleteInitialize was called, it sets the label for the user that was passed with UserId, otherwise sets the label for the user which is authenticated to the document.</para>
		/// <para>Docs: <see href="https://learn.microsoft.com/en-us/office/vba/api/office.sensitivitylabel.setlabel"/> </remarks>
		/// <param name="labelInfo">NetOffice.OfficeApi.LabelInfo labelInfo - The label information that needs to be set on the document.</param>
		/// <param name="context">object context - A caller defined context that can be returned with LabelChanged event to help ensure that the event was raised because of the SetLabel call.</param>
		[SupportByVersion("Office", 16)]
		public void SetLabel(NetOffice.OfficeApi.LabelInfo labelInfo, object context)
		{
			Factory.ExecuteMethod(this, "SetLabel", new object[]{ labelInfo, context });
		}

		#endregion

		#pragma warning restore
	}
}
