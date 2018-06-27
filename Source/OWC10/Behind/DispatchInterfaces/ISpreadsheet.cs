using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface ISpreadsheet 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class ISpreadsheet : COMObject, NetOffice.OWC10Api.ISpreadsheet
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
                    _contractType = typeof(NetOffice.OWC10Api.ISpreadsheet);
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
                    _type = typeof(ISpreadsheet);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ISpreadsheet() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public virtual NetOffice.OWC10Api._Range ActiveCell
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "ActiveCell");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Worksheet ActiveSheet
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Worksheet>(this, "ActiveSheet", typeof(NetOffice.OWC10Api.Worksheet));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Window ActiveWindow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Window>(this, "ActiveWindow", typeof(NetOffice.OWC10Api.Window));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Workbook ActiveWorkbook
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Workbook>(this, "ActiveWorkbook", typeof(NetOffice.OWC10Api.Workbook));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool AllowPropertyToolbox
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowPropertyToolbox");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowPropertyToolbox", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public virtual NetOffice.OWC10Api.ISpreadsheet Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api.ISpreadsheet>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool AutoFit
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AutoFit");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutoFit", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 Build
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Build");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string BuildNumber
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BuildNumber");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Enums.XlCalculation Calculation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.XlCalculation>(this, "Calculation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Calculation", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 CalculationVersion
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CalculationVersion");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool CanUndo
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "CanUndo");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
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
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public virtual NetOffice.OWC10Api._Range Cells
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Cells");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public virtual NetOffice.OWC10Api._Range Columns
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Columns");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.OCCommands Commands
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.OCCommands>(this, "Commands", typeof(NetOffice.OWC10Api.OCCommands));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Constants
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Constants");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string CSVData
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CSVData");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CSVData", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string CSVURL
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CSVURL");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CSVURL", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string DataMember
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DataMember");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DataMember", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.MSDATASRCApi.DataSource DataSource
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSDATASRCApi.DataSource>(this, "DataSource", typeof(NetOffice.MSDATASRCApi.DataSource));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "DataSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string DataType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DataType");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DataType", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool DesignMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DesignMode");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
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
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool DisplayBranding
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayBranding");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayBranding", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool DisplayColumnHeadings
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayColumnHeadings");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayColumnHeadings", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool DisplayDesignTimeUI
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayDesignTimeUI");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayDesignTimeUI", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool DisplayGridlines
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayGridlines");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayGridlines", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool DisplayHorizontalScrollBar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayHorizontalScrollBar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayHorizontalScrollBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool DisplayOfficeLogo
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayOfficeLogo");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayOfficeLogo", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool DisplayPropertyToolbox
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayPropertyToolbox");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayPropertyToolbox", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool DisplayRowHeadings
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayRowHeadings");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayRowHeadings", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool DisplayTitleBar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayTitleBar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayTitleBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool DisplayToolbar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayToolbar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayToolbar", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool DisplayVerticalScrollBar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayVerticalScrollBar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayVerticalScrollBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool DisplayWorkbookTabs
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayWorkbookTabs");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayWorkbookTabs", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool EnableEvents
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EnableEvents");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnableEvents", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool EnableUndo
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EnableUndo");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnableUndo", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string HTMLData
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "HTMLData");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HTMLData", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string HTMLURL
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "HTMLURL");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HTMLURL", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 InstanceID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "InstanceID");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object get_International(object index)
		{
			return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "International", index);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_International
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1), Redirect("get_International")]
		public virtual object International(object index)
		{
			return get_International(index);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.OWCLanguageSettings LanguageSettings
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.OWCLanguageSettings>(this, "LanguageSettings", typeof(NetOffice.OWC10Api.OWCLanguageSettings));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual object MaxHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "MaxHeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "MaxHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual object MaxWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "MaxWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "MaxWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 MajorVersion
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "MajorVersion");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string MinorVersion
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "MinorVersion");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool MoveAfterReturn
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MoveAfterReturn");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MoveAfterReturn", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Enums.XlDirection MoveAfterReturnDirection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.XlDirection>(this, "MoveAfterReturnDirection");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "MoveAfterReturnDirection", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Names Names
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Names>(this, "Names", typeof(NetOffice.OWC10Api.Names));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="cell1">object cell1</param>
		/// <param name="cell2">optional object cell2</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.OWC10Api._Range get_Range(object cell1, object cell2)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Range", typeof(NetOffice.OWC10Api._Range), cell1, cell2);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Range
		/// </summary>
		/// <param name="cell1">object cell1</param>
		/// <param name="cell2">optional object cell2</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Range")]
		public virtual NetOffice.OWC10Api._Range Range(object cell1, object cell2)
		{
			return get_Range(cell1, cell2);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="cell1">object cell1</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.OWC10Api._Range get_Range(object cell1)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Range", typeof(NetOffice.OWC10Api._Range), cell1);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Range
		/// </summary>
		/// <param name="cell1">object cell1</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Range")]
		public virtual NetOffice.OWC10Api._Range Range(object cell1)
		{
			return get_Range(cell1);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string RevisionNumber
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "RevisionNumber");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool RightToLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "RightToLeft");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RightToLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public virtual NetOffice.OWC10Api._Range Rows
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Rows");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool ScreenUpdating
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ScreenUpdating");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ScreenUpdating", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public virtual NetOffice.OWC10Api._Range Selection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Selection");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Sheets Sheets
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Sheets>(this, "Sheets", typeof(NetOffice.OWC10Api.Sheets));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.TitleBar TitleBar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.TitleBar>(this, "TitleBar", typeof(NetOffice.OWC10Api.TitleBar));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public virtual NetOffice.MSComctlLibApi.IToolbar Toolbar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSComctlLibApi.IToolbar>(this, "Toolbar");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string Value
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Value");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string Version
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Version");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string ViewableRange
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ViewableRange");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ViewableRange", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool ViewOnlyMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ViewOnlyMode");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Windows Windows
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Windows>(this, "Windows", typeof(NetOffice.OWC10Api.Windows));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Workbooks Workbooks
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Workbooks>(this, "Workbooks", typeof(NetOffice.OWC10Api.Workbooks));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Worksheets Worksheets
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Worksheets>(this, "Worksheets", typeof(NetOffice.OWC10Api.Worksheets));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string XMLData
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "XMLData");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "XMLData", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string XMLURL
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "XMLURL");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "XMLURL", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="addIn">object addIn</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void AddIn(object addIn)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddIn", addIn);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void BeginUndo()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "BeginUndo");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void Calculate()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Calculate");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void CalculateFull()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CalculateFull");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="cancel">optional bool Cancel = false</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void EndUndo(object cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EndUndo", cancel);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void EndUndo()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EndUndo");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="expression">object expression</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public virtual object _Evaluate(object expression)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_Evaluate", expression);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="expression">object expression</param>
		[SupportByVersion("OWC10", 1)]
		public virtual object Evaluate(object expression)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Evaluate", expression);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = </param>
		/// <param name="action">optional NetOffice.OWC10Api.Enums.SheetExportActionEnum Action = 1</param>
		/// <param name="format">optional NetOffice.OWC10Api.Enums.SheetExportFormat Format = 0</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void Export(object filename, object action, object format)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Export", filename, action, format);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void Export()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Export");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = </param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void Export(object filename)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Export", filename);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = </param>
		/// <param name="action">optional NetOffice.OWC10Api.Enums.SheetExportActionEnum Action = 1</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void Export(object filename, object action)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Export", filename, action);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public virtual void LocateDataSource()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LocateDataSource");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="bstr">string bstr</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public virtual object msDataSourceObject(string bstr)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "msDataSourceObject", bstr);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="range1">NetOffice.OWC10Api._Range range1</param>
		/// <param name="range2">NetOffice.OWC10Api._Range range2</param>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public virtual NetOffice.OWC10Api._Range RectIntersect(NetOffice.OWC10Api._Range range1, NetOffice.OWC10Api._Range range2)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "RectIntersect", range1, range2);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="range1">NetOffice.OWC10Api._Range range1</param>
		/// <param name="range2">NetOffice.OWC10Api._Range range2</param>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public virtual NetOffice.OWC10Api._Range RectUnion(NetOffice.OWC10Api._Range range1, NetOffice.OWC10Api._Range range2)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "RectUnion", range1, range2);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void Refresh()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Refresh");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void Repaint()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Repaint");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void ShowAbout()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowAbout");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="menu">object menu</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void ShowContextMenu(Int32 x, Int32 y, object menu)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowContextMenu", x, y, menu);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="topic">Int32 topic</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void ShowHelp(Int32 topic)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowHelp", topic);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void Undo()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Undo");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void UpdatePropertyToolbox()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UpdatePropertyToolbox");
		}

		#endregion

		#pragma warning restore
	}
}


