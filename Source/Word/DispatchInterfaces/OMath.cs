using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface OMath 
	/// SupportByVersion Word, 12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath"/> </remarks>
	[SupportByVersion("Word", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class OMath : COMObject
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
                    _type = typeof(OMath);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public OMath(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public OMath(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMath(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMath(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMath(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMath(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMath() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMath(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath.Application"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", NetOffice.WordApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath.Creator"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath.Parent"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath.Range"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Range Range
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "Range", NetOffice.WordApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath.Functions"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathFunctions Functions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathFunctions>(this, "Functions", NetOffice.WordApi.OMathFunctions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath.Type"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdOMathType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdOMathType>(this, "Type");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Type", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath.ParentOMath"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMath ParentOMath
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMath>(this, "ParentOMath", NetOffice.WordApi.OMath.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath.ParentFunction"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathFunction ParentFunction
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathFunction>(this, "ParentFunction", NetOffice.WordApi.OMathFunction.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath.ParentRow"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathMatRow ParentRow
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathMatRow>(this, "ParentRow", NetOffice.WordApi.OMathMatRow.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath.ParentCol"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathMatCol ParentCol
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathMatCol>(this, "ParentCol", NetOffice.WordApi.OMathMatCol.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath.ParentArg"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMath ParentArg
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMath>(this, "ParentArg", NetOffice.WordApi.OMath.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath.ArgIndex"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public Int32 ArgIndex
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ArgIndex");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath.NestingLevel"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public Int32 NestingLevel
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "NestingLevel");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath.ArgSize"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public Int32 ArgSize
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ArgSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ArgSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath.Breaks"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathBreaks Breaks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathBreaks>(this, "Breaks", NetOffice.WordApi.OMathBreaks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath.Justification"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdOMathJc Justification
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdOMathJc>(this, "Justification");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Justification", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath.AlignPoint"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public Int32 AlignPoint
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "AlignPoint");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AlignPoint", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath.Linearize"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void Linearize()
		{
			 Factory.ExecuteMethod(this, "Linearize");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath.BuildUp"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void BuildUp()
		{
			 Factory.ExecuteMethod(this, "BuildUp");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath.Remove"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void Remove()
		{
			 Factory.ExecuteMethod(this, "Remove");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath.ConvertToMathText"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ConvertToMathText()
		{
			 Factory.ExecuteMethod(this, "ConvertToMathText");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath.ConvertToNormalText"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ConvertToNormalText()
		{
			 Factory.ExecuteMethod(this, "ConvertToNormalText");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.OMath.ConvertToLiteralText"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ConvertToLiteralText()
		{
			 Factory.ExecuteMethod(this, "ConvertToLiteralText");
		}

		#endregion

		#pragma warning restore
	}
}
