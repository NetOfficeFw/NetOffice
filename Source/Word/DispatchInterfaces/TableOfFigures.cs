using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface TableOfFigures 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TableOfFigures"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class TableOfFigures : COMObject
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
                    _type = typeof(TableOfFigures);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public TableOfFigures(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public TableOfFigures(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableOfFigures(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableOfFigures(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableOfFigures(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableOfFigures(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableOfFigures() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableOfFigures(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TableOfFigures.Application"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", NetOffice.WordApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TableOfFigures.Creator"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TableOfFigures.Parent"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TableOfFigures.Caption"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TableOfFigures.IncludeLabel"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool IncludeLabel
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IncludeLabel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IncludeLabel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TableOfFigures.RightAlignPageNumbers"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool RightAlignPageNumbers
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RightAlignPageNumbers");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RightAlignPageNumbers", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TableOfFigures.UseHeadingStyles"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool UseHeadingStyles
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseHeadingStyles");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseHeadingStyles", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TableOfFigures.LowerHeadingLevel"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 LowerHeadingLevel
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "LowerHeadingLevel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LowerHeadingLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TableOfFigures.UpperHeadingLevel"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 UpperHeadingLevel
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "UpperHeadingLevel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UpperHeadingLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TableOfFigures.IncludePageNumbers"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool IncludePageNumbers
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IncludePageNumbers");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IncludePageNumbers", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TableOfFigures.Range"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Range
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "Range", NetOffice.WordApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TableOfFigures.UseFields"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool UseFields
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseFields");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseFields", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TableOfFigures.TableID"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string TableID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TableID");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TableID", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TableOfFigures.HeadingStyles"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.HeadingStyles HeadingStyles
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.HeadingStyles>(this, "HeadingStyles", NetOffice.WordApi.HeadingStyles.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TableOfFigures.TabLeader"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdTabLeader TabLeader
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdTabLeader>(this, "TabLeader");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TabLeader", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TableOfFigures.UseHyperlinks"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool UseHyperlinks
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseHyperlinks");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseHyperlinks", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TableOfFigures.HidePageNumbersInWeb"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool HidePageNumbersInWeb
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HidePageNumbersInWeb");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HidePageNumbersInWeb", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TableOfFigures.Delete"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TableOfFigures.UpdatePageNumbers"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void UpdatePageNumbers()
		{
			 Factory.ExecuteMethod(this, "UpdatePageNumbers");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TableOfFigures.Update"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Update()
		{
			 Factory.ExecuteMethod(this, "Update");
		}

		#endregion

		#pragma warning restore
	}
}
