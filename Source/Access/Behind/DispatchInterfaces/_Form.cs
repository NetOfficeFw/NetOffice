using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.AccessApi;

namespace NetOffice.AccessApi.Behind
{
	/// <summary>
	/// DispatchInterface _Form 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>	
	public class _Form : NetOffice.OfficeApi.Behind.IAccessible, NetOffice.AccessApi._Form
	{
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.AccessApi._Form);
                return _contractType;
            }
        }
        private static Type _contractType;


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

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _Form() : base()
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
		public virtual string FormName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FormName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FormName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821093.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string RecordSource
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "RecordSource");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RecordSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194672.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string Filter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Filter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Filter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195708.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool FilterOn
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FilterOn");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FilterOn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195510.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OrderBy
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OrderBy");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OrderBy", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197060.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool OrderByOn
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "OrderByOn");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OrderByOn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834341.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool AllowFilters
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowFilters");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowFilters", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193166.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string Caption
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Caption");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Caption", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822539.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte DefaultView
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "DefaultView");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192068.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte ViewsAllowed
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "ViewsAllowed");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ViewsAllowed", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool AllowEditing
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowEditing");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowEditing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int16 DefaultEditing
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "DefaultEditing");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultEditing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192851.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool AllowEdits
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowEdits");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowEdits", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821485.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool AllowDeletions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowDeletions");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowDeletions", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197373.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool AllowAdditions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowAdditions");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowAdditions", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845109.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool DataEntry
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DataEntry");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DataEntry", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual byte AllowUpdating
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "AllowUpdating");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowUpdating", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj249050.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte RecordsetType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "RecordsetType");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RecordsetType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197407.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte RecordLocks
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "RecordLocks");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RecordLocks", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834790.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte ScrollBars
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "ScrollBars");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ScrollBars", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196041.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool RecordSelectors
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "RecordSelectors");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RecordSelectors", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191795.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool NavigationButtons
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "NavigationButtons");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NavigationButtons", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836966.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool DividingLines
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DividingLines");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DividingLines", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194510.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool AutoResize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AutoResize");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutoResize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821162.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool AutoCenter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AutoCenter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutoCenter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845183.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool PopUp
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PopUp");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PopUp", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821033.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool Modal
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Modal");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Modal", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821190.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte BorderStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "BorderStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BorderStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823089.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool ControlBox
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ControlBox");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ControlBox", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool MinButton
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MinButton");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MinButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool MaxButton
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MaxButton");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MaxButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845417.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte MinMaxButtons
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "MinMaxButtons");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MinMaxButtons", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823184.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool CloseButton
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "CloseButton");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CloseButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool WhatsThisButton
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "WhatsThisButton");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "WhatsThisButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192847.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 Width
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Width");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Width", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193484.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string Picture
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Picture");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Picture", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197672.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte PictureType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "PictureType");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PictureType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822034.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte PictureSizeMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "PictureSizeMode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PictureSizeMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197378.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte PictureAlignment
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "PictureAlignment");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PictureAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197664.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool PictureTiling
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PictureTiling");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PictureTiling", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194916.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte Cycle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "Cycle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Cycle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822480.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string MenuBar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "MenuBar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MenuBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820738.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string Toolbar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Toolbar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Toolbar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836305.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool ShortcutMenu
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShortcutMenu");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShortcutMenu", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822064.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string ShortcutMenuBar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ShortcutMenuBar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShortcutMenuBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836275.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 GridX
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "GridX");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835066.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 GridY
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "GridY");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837245.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool LayoutForPrint
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "LayoutForPrint");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LayoutForPrint", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821174.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool FastLaserPrinting
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FastLaserPrinting");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FastLaserPrinting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195832.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string HelpFile
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "HelpFile");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HelpFile", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845889.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 HelpContextId
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "HelpContextId");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HelpContextId", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845493.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 RowHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "RowHeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RowHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835977.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string DatasheetFontName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DatasheetFontName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DatasheetFontName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194592.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 DatasheetFontHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "DatasheetFontHeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DatasheetFontHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195549.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 DatasheetFontWeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "DatasheetFontWeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DatasheetFontWeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192317.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool DatasheetFontItalic
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DatasheetFontItalic");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DatasheetFontItalic", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820963.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool DatasheetFontUnderline
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DatasheetFontUnderline");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DatasheetFontUnderline", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual byte TabularCharSet
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "TabularCharSet");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TabularCharSet", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195269.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte DatasheetGridlinesBehavior
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "DatasheetGridlinesBehavior");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DatasheetGridlinesBehavior", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197658.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 DatasheetGridlinesColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "DatasheetGridlinesColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DatasheetGridlinesColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192508.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte DatasheetCellsEffect
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "DatasheetCellsEffect");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DatasheetCellsEffect", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197788.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 DatasheetForeColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "DatasheetForeColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DatasheetForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool ShowGrid
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowGrid");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowGrid", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195276.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 DatasheetBackColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "DatasheetBackColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DatasheetBackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197072.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 Hwnd
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Hwnd");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Hwnd", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835632.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Count");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Count", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845127.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 Page
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Page");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Page", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197679.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 Pages
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Pages");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Pages", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 LogicalPageWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "LogicalPageWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LogicalPageWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 LogicalPageHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "LogicalPageHeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LogicalPageHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 ZoomControl
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ZoomControl");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ZoomControl", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196785.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool Visible
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Visible");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Visible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195590.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool Painting
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Painting");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Painting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845141.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object PrtMip
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PrtMip");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "PrtMip", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820951.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object PrtDevMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PrtDevMode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "PrtDevMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845154.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object PrtDevNames
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PrtDevNames");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "PrtDevNames", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194534.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 FrozenColumns
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "FrozenColumns");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FrozenColumns", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835682.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Bookmark
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Bookmark");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Bookmark", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual byte TabularFamily
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "TabularFamily");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TabularFamily", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string _Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "_Name");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "_Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197620.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string PaletteSource
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PaletteSource");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PaletteSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196472.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string Tag
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Tag");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Tag", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845517.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object PaintPalette
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PaintPalette");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "PaintPalette", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnMenu
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnMenu");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnMenu", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836583.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object OpenArgs
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "OpenArgs");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "OpenArgs", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int16 ConnectSynch
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ConnectSynch");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ConnectSynch", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822706.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnCurrent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnCurrent");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnCurrent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191901.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnInsert
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnInsert");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnInsert", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194954.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string BeforeInsert
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BeforeInsert");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BeforeInsert", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197713.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string AfterInsert
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AfterInsert");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AfterInsert", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822073.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string BeforeUpdate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BeforeUpdate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BeforeUpdate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193798.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string AfterUpdate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AfterUpdate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AfterUpdate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835673.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnDirty
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnDirty");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnDirty", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197932.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnDelete
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnDelete");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnDelete", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197065.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string BeforeDelConfirm
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BeforeDelConfirm");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BeforeDelConfirm", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837215.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string AfterDelConfirm
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AfterDelConfirm");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AfterDelConfirm", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845483.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnOpen
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnOpen");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnOpen", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196803.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnLoad
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnLoad");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnLoad", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196776.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnResize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnResize");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnResize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195730.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnUnload
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnUnload");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnUnload", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821764.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnClose
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnClose");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnClose", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821479.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnActivate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnActivate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnActivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823016.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnDeactivate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnDeactivate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnDeactivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844857.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnGotFocus
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnGotFocus");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnGotFocus", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192085.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnLostFocus
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnLostFocus");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnLostFocus", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820744.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnClick
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnClick");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnClick", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197372.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnDblClick
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnDblClick");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnDblClick", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834386.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnMouseDown
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnMouseDown");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnMouseDown", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822839.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnMouseMove
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnMouseMove");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnMouseMove", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197110.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnMouseUp
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnMouseUp");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnMouseUp", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195851.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnKeyDown
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnKeyDown");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnKeyDown", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845625.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnKeyUp
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnKeyUp");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnKeyUp", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845717.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnKeyPress
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnKeyPress");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnKeyPress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836983.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool KeyPreview
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "KeyPreview");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "KeyPreview", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836950.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnError
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnError");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnError", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193563.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnFilter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnFilter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194649.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnApplyFilter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnApplyFilter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnApplyFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821383.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnTimer
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnTimer");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnTimer", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836371.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 TimerInterval
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "TimerInterval");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TimerInterval", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194309.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool Dirty
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Dirty");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Dirty", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196494.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 WindowWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "WindowWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "WindowWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194007.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 WindowHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "WindowHeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "WindowHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834753.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 CurrentView
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "CurrentView");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CurrentView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835055.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 CurrentSectionTop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "CurrentSectionTop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CurrentSectionTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194568.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 CurrentSectionLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "CurrentSectionLeft");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CurrentSectionLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835384.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 SelLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SelLeft");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SelLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194148.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 SelTop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SelTop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SelTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821151.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 SelWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SelWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SelWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823187.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 SelHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SelHeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SelHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821182.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 CurrentRecord
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CurrentRecord");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CurrentRecord", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845061.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object PictureData
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PictureData");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "PictureData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196178.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 InsideHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "InsideHeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "InsideHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834321.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 InsideWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "InsideWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "InsideWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193513.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object PicturePalette
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PicturePalette");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "PicturePalette", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822494.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool HasModule
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasModule");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HasModule", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 acHiddenCurrentPage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "acHiddenCurrentPage");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "acHiddenCurrentPage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191871.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte Orientation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "Orientation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Orientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool AllowDesignChanges
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowDesignChanges");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowDesignChanges", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845592.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string ServerFilter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ServerFilter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ServerFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837027.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool ServerFilterByForm
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ServerFilterByForm");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ServerFilterByForm", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845727.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 MaxRecords
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "MaxRecords");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MaxRecords", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191879.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string UniqueTable
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "UniqueTable");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UniqueTable", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845228.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string ResyncCommand
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ResyncCommand");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ResyncCommand", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837198.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string InputParameters
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "InputParameters");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "InputParameters", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195580.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool MaxRecButton
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MaxRecButton");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MaxRecButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835430.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Application>(this, "Application", typeof(NetOffice.AccessApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834458.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198278.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 NewRecord
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "NewRecord");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845144.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Control ActiveControl
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Control>(this, "ActiveControl", typeof(NetOffice.AccessApi.Control));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836869.aspx </remarks>
		/// <param name="controlType">Int32 controlType</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.AccessApi.Control get_DefaultControl(Int32 controlType)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Control>(this, "DefaultControl", typeof(NetOffice.AccessApi.Control), controlType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_DefaultControl
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836869.aspx </remarks>
		/// <param name="controlType">Int32 controlType</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_DefaultControl")]
		public virtual NetOffice.AccessApi.Control DefaultControl(Int32 controlType)
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
		public virtual object Dynaset
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Dynaset");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835062.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object RecordsetClone
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "RecordsetClone");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822528.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Recordset
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Recordset");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Recordset", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835642.aspx </remarks>
		/// <param name="index">object index</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.AccessApi.Section get_Section(object index)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Section>(this, "Section", typeof(NetOffice.AccessApi.Section), index);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Section
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835642.aspx </remarks>
		/// <param name="index">object index</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_Section")]
		public virtual NetOffice.AccessApi.Section Section(object index)
		{
			return get_Section(index);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194652.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Form Form
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Form>(this, "Form", typeof(NetOffice.AccessApi.Form));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836688.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Module Module
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Module>(this, "Module", typeof(NetOffice.AccessApi.Module));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194921.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Properties Properties
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Properties>(this, "Properties", typeof(NetOffice.AccessApi.Properties));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.AccessApi.Control ConnectControl
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Control>(this, "ConnectControl", typeof(NetOffice.AccessApi.Control));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845021.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Controls Controls
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Controls>(this, "Controls", typeof(NetOffice.AccessApi.Controls));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192050.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845216.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 SubdatasheetHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "SubdatasheetHeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SubdatasheetHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194094.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool SubdatasheetExpanded
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SubdatasheetExpanded");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SubdatasheetExpanded", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195175.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Undo()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Undo");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194887.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Recalc()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Recalc");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191903.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Requery()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Requery");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836021.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Refresh()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Refresh");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834494.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Repaint()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Repaint");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197662.aspx </remarks>
		/// <param name="pageNumber">Int32 pageNumber</param>
		/// <param name="right">optional Int32 Right = 0</param>
		/// <param name="down">optional Int32 Down = 0</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void GoToPage(Int32 pageNumber, object right, object down)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "GoToPage", pageNumber, right, down);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197662.aspx </remarks>
		/// <param name="pageNumber">Int32 pageNumber</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void GoToPage(Int32 pageNumber)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "GoToPage", pageNumber);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197662.aspx </remarks>
		/// <param name="pageNumber">Int32 pageNumber</param>
		/// <param name="right">optional Int32 Right = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void GoToPage(Int32 pageNumber, object right)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "GoToPage", pageNumber, right);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821776.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void SetFocus()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetFocus");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrExpr">string bstrExpr</param>
		/// <param name="ppsa">optional object[] ppsa</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object _Evaluate(string bstrExpr, object[] ppsa)
		{
            object[] paramsArray = Invoker.ValidateParamsArray(bstrExpr, (object)ppsa);
            object returnItem = Invoker.MethodReturn(this, "_Evaluate", paramsArray);
            if ((null != returnItem) && (returnItem is MarshalByRefObject))
            {
                ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem, true);
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
		public virtual object _Evaluate(string bstrExpr)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_Evaluate", bstrExpr);
		}

        #endregion

        #region IEnumerableProvider<object>

        ICOMObject IEnumerableProvider<object>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsMethod(parent, this, false);
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
        public virtual IEnumerator<object> GetEnumerator()
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
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this, true);
		}

		#endregion

		#pragma warning restore
	}
}

