using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface List 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192789.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class List : COMObject
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
                    _type = typeof(List);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public List(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public List(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public List(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public List(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public List(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public List(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public List() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public List(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840272.aspx </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193376.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ListParagraphs ListParagraphs
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ListParagraphs>(this, "ListParagraphs", NetOffice.WordApi.ListParagraphs.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838053.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool SingleListTemplate
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SingleListTemplate");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194676.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836145.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194035.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194057.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public string StyleName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "StyleName");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193093.aspx </remarks>
		/// <param name="numberType">optional object numberType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ConvertNumbersToText(object numberType)
		{
			 Factory.ExecuteMethod(this, "ConvertNumbersToText", numberType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193093.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ConvertNumbersToText()
		{
			 Factory.ExecuteMethod(this, "ConvertNumbersToText");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838078.aspx </remarks>
		/// <param name="numberType">optional object numberType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void RemoveNumbers(object numberType)
		{
			 Factory.ExecuteMethod(this, "RemoveNumbers", numberType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838078.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void RemoveNumbers()
		{
			 Factory.ExecuteMethod(this, "RemoveNumbers");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820795.aspx </remarks>
		/// <param name="numberType">optional object numberType</param>
		/// <param name="level">optional object level</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 CountNumberedItems(object numberType, object level)
		{
			return Factory.ExecuteInt32MethodGet(this, "CountNumberedItems", numberType, level);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820795.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 CountNumberedItems()
		{
			return Factory.ExecuteInt32MethodGet(this, "CountNumberedItems");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820795.aspx </remarks>
		/// <param name="numberType">optional object numberType</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 CountNumberedItems(object numberType)
		{
			return Factory.ExecuteInt32MethodGet(this, "CountNumberedItems", numberType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ApplyListTemplateOld(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList)
		{
			 Factory.ExecuteMethod(this, "ApplyListTemplateOld", listTemplate, continuePreviousList);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ApplyListTemplateOld(NetOffice.WordApi.ListTemplate listTemplate)
		{
			 Factory.ExecuteMethod(this, "ApplyListTemplateOld", listTemplate);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196826.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdContinue CanContinuePreviousList(NetOffice.WordApi.ListTemplate listTemplate)
		{
			return Factory.ExecuteEnumMethodGet<NetOffice.WordApi.Enums.WdContinue>(this, "CanContinuePreviousList", listTemplate);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196090.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		/// <param name="defaultListBehavior">optional object defaultListBehavior</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ApplyListTemplate(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object defaultListBehavior)
		{
			 Factory.ExecuteMethod(this, "ApplyListTemplate", listTemplate, continuePreviousList, defaultListBehavior);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196090.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ApplyListTemplate(NetOffice.WordApi.ListTemplate listTemplate)
		{
			 Factory.ExecuteMethod(this, "ApplyListTemplate", listTemplate);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196090.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ApplyListTemplate(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList)
		{
			 Factory.ExecuteMethod(this, "ApplyListTemplate", listTemplate, continuePreviousList);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191850.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		/// <param name="defaultListBehavior">optional object defaultListBehavior</param>
		/// <param name="applyLevel">optional object applyLevel</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ApplyListTemplateWithLevel(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object defaultListBehavior, object applyLevel)
		{
			 Factory.ExecuteMethod(this, "ApplyListTemplateWithLevel", listTemplate, continuePreviousList, defaultListBehavior, applyLevel);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191850.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ApplyListTemplateWithLevel(NetOffice.WordApi.ListTemplate listTemplate)
		{
			 Factory.ExecuteMethod(this, "ApplyListTemplateWithLevel", listTemplate);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191850.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ApplyListTemplateWithLevel(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList)
		{
			 Factory.ExecuteMethod(this, "ApplyListTemplateWithLevel", listTemplate, continuePreviousList);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191850.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		/// <param name="defaultListBehavior">optional object defaultListBehavior</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ApplyListTemplateWithLevel(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object defaultListBehavior)
		{
			 Factory.ExecuteMethod(this, "ApplyListTemplateWithLevel", listTemplate, continuePreviousList, defaultListBehavior);
		}

		#endregion

		#pragma warning restore
	}
}
