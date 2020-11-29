using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface TextFrame 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class TextFrame : COMObject
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
                    _type = typeof(TextFrame);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public TextFrame(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public TextFrame(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextFrame(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextFrame(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextFrame(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextFrame(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextFrame() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextFrame(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.Creator"/> </remarks>
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
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.Parent"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape Parent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Shape>(this, "Parent", NetOffice.WordApi.Shape.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.MarginBottom"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single MarginBottom
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "MarginBottom");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MarginBottom", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.MarginLeft"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single MarginLeft
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "MarginLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MarginLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.MarginRight"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single MarginRight
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "MarginRight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MarginRight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.MarginTop"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single MarginTop
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "MarginTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MarginTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.Orientation"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTextOrientation Orientation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTextOrientation>(this, "Orientation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Orientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.TextRange"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range TextRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "TextRange", NetOffice.WordApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.ContainingRange"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range ContainingRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "ContainingRange", NetOffice.WordApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.Next"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TextFrame Next
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.TextFrame>(this, "Next", NetOffice.WordApi.TextFrame.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Next", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.Previous"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TextFrame Previous
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.TextFrame>(this, "Previous", NetOffice.WordApi.TextFrame.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Previous", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.Overflowing"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Overflowing
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Overflowing");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.HasText"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 HasText
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "HasText");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.AutoSize"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public Int32 AutoSize
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "AutoSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.WordWrap"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public Int32 WordWrap
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "WordWrap");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WordWrap", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.VerticalAnchor"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoVerticalAnchor VerticalAnchor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoVerticalAnchor>(this, "VerticalAnchor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "VerticalAnchor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.HorizontalAnchor"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoHorizontalAnchor HorizontalAnchor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoHorizontalAnchor>(this, "HorizontalAnchor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "HorizontalAnchor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.PathFormat"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoPathFormat PathFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPathFormat>(this, "PathFormat");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PathFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.WarpFormat"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoWarpFormat WarpFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoWarpFormat>(this, "WarpFormat");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "WarpFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.Column"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.OfficeApi.TextColumn2 Column
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextColumn2>(this, "Column", NetOffice.OfficeApi.TextColumn2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.ThreeD"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.ThreeDFormat ThreeD
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ThreeDFormat>(this, "ThreeD", NetOffice.WordApi.ThreeDFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.NoTextRotation"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState NoTextRotation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "NoTextRotation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "NoTextRotation", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.BreakForwardLink"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void BreakForwardLink()
		{
			 Factory.ExecuteMethod(this, "BreakForwardLink");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.ValidLinkTarget"/> </remarks>
		/// <param name="targetTextFrame">NetOffice.WordApi.TextFrame targetTextFrame</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ValidLinkTarget(NetOffice.WordApi.TextFrame targetTextFrame)
		{
			return Factory.ExecuteBoolMethodGet(this, "ValidLinkTarget", targetTextFrame);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TextFrame.DeleteText"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public void DeleteText()
		{
			 Factory.ExecuteMethod(this, "DeleteText");
		}

		#endregion

		#pragma warning restore
	}
}
