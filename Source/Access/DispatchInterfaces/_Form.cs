using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _Form 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Method)]
	public class _Form : NetOffice.OfficeApi.IAccessible, IEnumerableProvider<object>
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
                    _type = typeof(_Form);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Form(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Form(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Form(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Form(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Form(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Form(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Form() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Form(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string FormName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FormName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FormName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.RecordSource"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string RecordSource
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RecordSource");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecordSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Filter(property)"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string Filter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Filter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Filter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.FilterOn"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool FilterOn
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FilterOn");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FilterOn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OrderBy"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OrderBy
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OrderBy");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OrderBy", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OrderByOn"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool OrderByOn
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "OrderByOn");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OrderByOn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.AllowFilters"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool AllowFilters
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowFilters");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowFilters", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Caption"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string Caption
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Caption");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Caption", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.DefaultView"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte DefaultView
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "DefaultView");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.ViewsAllowed"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte ViewsAllowed
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "ViewsAllowed");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ViewsAllowed", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool AllowEditing
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowEditing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowEditing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 DefaultEditing
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "DefaultEditing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultEditing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.AllowEdits"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool AllowEdits
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowEdits");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowEdits", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.AllowDeletions"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool AllowDeletions
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowDeletions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowDeletions", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.AllowAdditions"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool AllowAdditions
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowAdditions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowAdditions", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.DataEntry"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool DataEntry
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DataEntry");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataEntry", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public byte AllowUpdating
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "AllowUpdating");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowUpdating", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.form.recordsettype"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte RecordsetType
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "RecordsetType");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecordsetType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.RecordLocks"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte RecordLocks
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "RecordLocks");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecordLocks", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.ScrollBars"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte ScrollBars
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "ScrollBars");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ScrollBars", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.RecordSelectors"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool RecordSelectors
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RecordSelectors");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecordSelectors", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.NavigationButtons"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool NavigationButtons
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "NavigationButtons");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NavigationButtons", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.DividingLines"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool DividingLines
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DividingLines");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DividingLines", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.AutoResize"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool AutoResize
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoResize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoResize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.AutoCenter"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool AutoCenter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoCenter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoCenter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.PopUp"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool PopUp
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PopUp");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PopUp", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Modal"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool Modal
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Modal");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Modal", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.BorderStyle"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte BorderStyle
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "BorderStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BorderStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.ControlBox"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool ControlBox
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ControlBox");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ControlBox", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool MinButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MinButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MinButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool MaxButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MaxButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MaxButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.MinMaxButtons"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte MinMaxButtons
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "MinMaxButtons");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MinMaxButtons", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.CloseButton"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool CloseButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CloseButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CloseButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool WhatsThisButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "WhatsThisButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WhatsThisButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Width"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 Width
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Width");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Width", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Picture"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string Picture
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Picture");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Picture", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.PictureType"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte PictureType
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "PictureType");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.PictureSizeMode"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte PictureSizeMode
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "PictureSizeMode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureSizeMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.PictureAlignment"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte PictureAlignment
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "PictureAlignment");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.PictureTiling"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool PictureTiling
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PictureTiling");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureTiling", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Cycle"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte Cycle
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "Cycle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Cycle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.MenuBar"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string MenuBar
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MenuBar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MenuBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Toolbar"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string Toolbar
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Toolbar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Toolbar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.ShortcutMenu"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool ShortcutMenu
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShortcutMenu");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShortcutMenu", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.ShortcutMenuBar"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string ShortcutMenuBar
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ShortcutMenuBar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShortcutMenuBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.GridX"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 GridX
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "GridX");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.GridY"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 GridY
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "GridY");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.LayoutForPrint"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool LayoutForPrint
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LayoutForPrint");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LayoutForPrint", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.FastLaserPrinting"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool FastLaserPrinting
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FastLaserPrinting");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FastLaserPrinting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.HelpFile"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string HelpFile
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HelpFile");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HelpFile", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.HelpContextId"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 HelpContextId
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "HelpContextId");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HelpContextId", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.RowHeight"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 RowHeight
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "RowHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RowHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.DatasheetFontName"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string DatasheetFontName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DatasheetFontName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetFontName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.DatasheetFontHeight"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 DatasheetFontHeight
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "DatasheetFontHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetFontHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.DatasheetFontWeight"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 DatasheetFontWeight
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "DatasheetFontWeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetFontWeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.DatasheetFontItalic"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool DatasheetFontItalic
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DatasheetFontItalic");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetFontItalic", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.DatasheetFontUnderline"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool DatasheetFontUnderline
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DatasheetFontUnderline");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetFontUnderline", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public byte TabularCharSet
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "TabularCharSet");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TabularCharSet", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.DatasheetGridlinesBehavior"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte DatasheetGridlinesBehavior
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "DatasheetGridlinesBehavior");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetGridlinesBehavior", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.DatasheetGridlinesColor"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 DatasheetGridlinesColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DatasheetGridlinesColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetGridlinesColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.DatasheetCellsEffect"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte DatasheetCellsEffect
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "DatasheetCellsEffect");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetCellsEffect", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.DatasheetForeColor"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 DatasheetForeColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DatasheetForeColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool ShowGrid
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowGrid");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowGrid", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.DatasheetBackColor"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 DatasheetBackColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DatasheetBackColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetBackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Hwnd"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 Hwnd
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Hwnd");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Hwnd", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Count"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 Count
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Count");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Count", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Page"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 Page
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Page");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Page", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Pages"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 Pages
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Pages");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Pages", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 LogicalPageWidth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "LogicalPageWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LogicalPageWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 LogicalPageHeight
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "LogicalPageHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LogicalPageHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 ZoomControl
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ZoomControl");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ZoomControl", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Visible"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool Visible
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Visible");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Visible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Painting"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool Painting
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Painting");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Painting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.PrtMip"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object PrtMip
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PrtMip");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PrtMip", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.PrtDevMode"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object PrtDevMode
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PrtDevMode");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PrtDevMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.PrtDevNames"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object PrtDevNames
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PrtDevNames");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PrtDevNames", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.FrozenColumns"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 FrozenColumns
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "FrozenColumns");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FrozenColumns", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Bookmark"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object Bookmark
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Bookmark");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Bookmark", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public byte TabularFamily
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "TabularFamily");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TabularFamily", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string _Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "_Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "_Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.PaletteSource"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string PaletteSource
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PaletteSource");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PaletteSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Tag"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string Tag
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Tag");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Tag", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.PaintPalette"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object PaintPalette
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PaintPalette");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PaintPalette", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnMenu
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnMenu");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnMenu", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OpenArgs"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object OpenArgs
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "OpenArgs");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "OpenArgs", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 ConnectSynch
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ConnectSynch");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ConnectSynch", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnCurrent"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnCurrent
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnCurrent");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnCurrent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnInsert"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnInsert
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnInsert");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnInsert", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.BeforeInsert(property)"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string BeforeInsert
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeInsert");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeInsert", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.AfterInsert(property)"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string AfterInsert
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterInsert");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterInsert", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.BeforeUpdate-property"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string BeforeUpdate
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeUpdate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeUpdate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.AfterUpdate-property"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string AfterUpdate
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterUpdate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterUpdate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnDirty"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnDirty
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnDirty");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnDirty", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnDelete"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnDelete
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnDelete");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnDelete", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.BeforeDelConfirm(property)"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string BeforeDelConfirm
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeDelConfirm");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeDelConfirm", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.AfterDelConfirm(property)"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string AfterDelConfirm
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterDelConfirm");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterDelConfirm", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnOpen"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnOpen
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnOpen");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnOpen", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnLoad"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnLoad
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnLoad");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnLoad", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnResize"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnResize
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnResize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnResize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnUnload"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnUnload
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnUnload");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnUnload", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnClose"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnClose
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnClose");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnClose", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnActivate"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnActivate
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnActivate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnActivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnDeactivate"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnDeactivate
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnDeactivate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnDeactivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnGotFocus"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnGotFocus
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnGotFocus");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnGotFocus", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnLostFocus"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnLostFocus
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnLostFocus");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnLostFocus", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnClick"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnClick
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnClick");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnClick", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnDblClick"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnDblClick
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnDblClick");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnDblClick", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnMouseDown"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnMouseDown
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnMouseDown");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnMouseDown", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnMouseMove"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnMouseMove
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnMouseMove");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnMouseMove", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnMouseUp"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnMouseUp
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnMouseUp");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnMouseUp", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnKeyDown"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnKeyDown
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnKeyDown");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnKeyDown", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnKeyUp"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnKeyUp
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnKeyUp");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnKeyUp", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnKeyPress"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnKeyPress
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnKeyPress");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnKeyPress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.KeyPreview"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool KeyPreview
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "KeyPreview");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "KeyPreview", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnError"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnError
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnError");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnError", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnFilter"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnFilter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnFilter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnApplyFilter"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnApplyFilter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnApplyFilter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnApplyFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.OnTimer"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnTimer
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnTimer");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnTimer", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.TimerInterval"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 TimerInterval
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TimerInterval");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TimerInterval", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Dirty(property)"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool Dirty
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Dirty");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Dirty", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.WindowWidth"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 WindowWidth
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "WindowWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WindowWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.WindowHeight"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 WindowHeight
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "WindowHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WindowHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.CurrentView"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 CurrentView
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "CurrentView");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CurrentView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.CurrentSectionTop"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 CurrentSectionTop
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "CurrentSectionTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CurrentSectionTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.CurrentSectionLeft"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 CurrentSectionLeft
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "CurrentSectionLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CurrentSectionLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.SelLeft"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 SelLeft
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SelLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SelLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.SelTop"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 SelTop
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SelTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SelTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.SelWidth"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 SelWidth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SelWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SelWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.SelHeight"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 SelHeight
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SelHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SelHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.CurrentRecord"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 CurrentRecord
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CurrentRecord");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CurrentRecord", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.PictureData"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object PictureData
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PictureData");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PictureData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.InsideHeight"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 InsideHeight
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "InsideHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InsideHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.InsideWidth"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 InsideWidth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "InsideWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InsideWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.PicturePalette"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object PicturePalette
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PicturePalette");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PicturePalette", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.HasModule"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool HasModule
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasModule");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasModule", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 acHiddenCurrentPage
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "acHiddenCurrentPage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "acHiddenCurrentPage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Orientation"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte Orientation
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "Orientation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Orientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool AllowDesignChanges
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowDesignChanges");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowDesignChanges", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.ServerFilter"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string ServerFilter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ServerFilter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ServerFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.ServerFilterByForm"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool ServerFilterByForm
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ServerFilterByForm");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ServerFilterByForm", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.MaxRecords"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 MaxRecords
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MaxRecords");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MaxRecords", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.UniqueTable"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string UniqueTable
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "UniqueTable");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UniqueTable", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.ResyncCommand"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string ResyncCommand
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ResyncCommand");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ResyncCommand", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.InputParameters"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string InputParameters
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "InputParameters");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InputParameters", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.MaxRecButton"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool MaxRecButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MaxRecButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MaxRecButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Application"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Application>(this, "Application", NetOffice.AccessApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Parent"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.NewRecord"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 NewRecord
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "NewRecord");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.ActiveControl"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Control ActiveControl
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Control>(this, "ActiveControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.DefaultControl"/> </remarks>
		/// <param name="controlType">Int32 controlType</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.AccessApi.Control get_DefaultControl(Int32 controlType)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Control>(this, "DefaultControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType, controlType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_DefaultControl
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.DefaultControl"/> </remarks>
		/// <param name="controlType">Int32 controlType</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_DefaultControl")]
		public NetOffice.AccessApi.Control DefaultControl(Int32 controlType)
		{
			return get_DefaultControl(controlType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Dynaset
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Dynaset");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.RecordsetClone"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		public object RecordsetClone
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "RecordsetClone");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Recordset"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		public object Recordset
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Recordset");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Recordset", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Section"/> </remarks>
		/// <param name="index">object index</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.AccessApi.Section get_Section(object index)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Section>(this, "Section", NetOffice.AccessApi.Section.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Section
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Section"/> </remarks>
		/// <param name="index">object index</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_Section")]
		public NetOffice.AccessApi.Section Section(object index)
		{
			return get_Section(index);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Form"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Form Form
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Form>(this, "Form", NetOffice.AccessApi.Form.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Module"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Module Module
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Module>(this, "Module", NetOffice.AccessApi.Module.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Properties"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Properties Properties
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Properties>(this, "Properties", NetOffice.AccessApi.Properties.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.AccessApi.Control ConnectControl
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Control>(this, "ConnectControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Controls"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Controls Controls
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Controls>(this, "Controls", NetOffice.AccessApi.Controls.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Name"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.SubdatasheetHeight"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 SubdatasheetHeight
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "SubdatasheetHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SubdatasheetHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.SubdatasheetExpanded"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool SubdatasheetExpanded
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SubdatasheetExpanded");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SubdatasheetExpanded", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Undo(method)"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void Undo()
		{
			 Factory.ExecuteMethod(this, "Undo");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Recalc"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void Recalc()
		{
			 Factory.ExecuteMethod(this, "Recalc");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Requery"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void Requery()
		{
			 Factory.ExecuteMethod(this, "Requery");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Refresh"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void Refresh()
		{
			 Factory.ExecuteMethod(this, "Refresh");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.Repaint"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void Repaint()
		{
			 Factory.ExecuteMethod(this, "Repaint");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.GoToPage"/> </remarks>
		/// <param name="pageNumber">Int32 pageNumber</param>
		/// <param name="right">optional Int32 Right = 0</param>
		/// <param name="down">optional Int32 Down = 0</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void GoToPage(Int32 pageNumber, object right, object down)
		{
			 Factory.ExecuteMethod(this, "GoToPage", pageNumber, right, down);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.GoToPage"/> </remarks>
		/// <param name="pageNumber">Int32 pageNumber</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void GoToPage(Int32 pageNumber)
		{
			 Factory.ExecuteMethod(this, "GoToPage", pageNumber);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.GoToPage"/> </remarks>
		/// <param name="pageNumber">Int32 pageNumber</param>
		/// <param name="right">optional Int32 Right = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void GoToPage(Int32 pageNumber, object right)
		{
			 Factory.ExecuteMethod(this, "GoToPage", pageNumber, right);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Form.SetFocus"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void SetFocus()
		{
			 Factory.ExecuteMethod(this, "SetFocus");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrExpr">string bstrExpr</param>
		/// <param name="ppsa">optional object[] ppsa</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object _Evaluate(string bstrExpr, object[] ppsa)
		{
            object[] paramsArray = Invoker.ValidateParamsArray(bstrExpr, (object)ppsa);
            object returnItem = Invoker.MethodReturn(this, "_Evaluate", paramsArray);
            if ((null != returnItem) && (returnItem is MarshalByRefObject))
            {
                ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
                return newObject;
            }
            else
            {
                return returnItem;
            }
        }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrExpr">string bstrExpr</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object _Evaluate(string bstrExpr)
		{
			return Factory.ExecuteVariantMethodGet(this, "_Evaluate", bstrExpr);
		}

        #endregion

        #region IEnumerableProvider<object>

        ICOMObject IEnumerableProvider<object>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsMethod(parent, this);
        }

        IEnumerable IEnumerableProvider<object>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, true);
        }

        #endregion

        #region IEnumerable<object>

        /// <summary>
        /// SupportByVersion Access, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public IEnumerator<object> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (object item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Access, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this);
		}

		#endregion

		#pragma warning restore
	}
}