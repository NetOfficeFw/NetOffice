using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface ListFormat 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820923.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ListFormat : COMObject
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
                    _type = typeof(ListFormat);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ListFormat(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ListFormat(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ListFormat(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ListFormat(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ListFormat(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ListFormat(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ListFormat() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ListFormat(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844779.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 ListLevelNumber
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ListLevelNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ListLevelNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839514.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.List List
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.List>(this, "List", NetOffice.WordApi.List.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821156.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ListTemplate ListTemplate
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ListTemplate>(this, "ListTemplate", NetOffice.WordApi.ListTemplate.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196311.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 ListValue
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ListValue");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836742.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool SingleList
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SingleList");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835216.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197765.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdListType ListType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdListType>(this, "ListType");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836909.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string ListString
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ListString");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194491.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195720.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194230.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837259.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape ListPictureBullet
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.InlineShape>(this, "ListPictureBullet", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196639.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdContinue CanContinuePreviousList(NetOffice.WordApi.ListTemplate listTemplate)
		{
			return Factory.ExecuteEnumMethodGet<NetOffice.WordApi.Enums.WdContinue>(this, "CanContinuePreviousList", listTemplate);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821874.aspx </remarks>
		/// <param name="numberType">optional object numberType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void RemoveNumbers(object numberType)
		{
			 Factory.ExecuteMethod(this, "RemoveNumbers", numberType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821874.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void RemoveNumbers()
		{
			 Factory.ExecuteMethod(this, "RemoveNumbers");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196580.aspx </remarks>
		/// <param name="numberType">optional object numberType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ConvertNumbersToText(object numberType)
		{
			 Factory.ExecuteMethod(this, "ConvertNumbersToText", numberType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196580.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ConvertNumbersToText()
		{
			 Factory.ExecuteMethod(this, "ConvertNumbersToText");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820721.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820721.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 CountNumberedItems()
		{
			return Factory.ExecuteInt32MethodGet(this, "CountNumberedItems");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820721.aspx </remarks>
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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ApplyBulletDefaultOld()
		{
			 Factory.ExecuteMethod(this, "ApplyBulletDefaultOld");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ApplyNumberDefaultOld()
		{
			 Factory.ExecuteMethod(this, "ApplyNumberDefaultOld");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ApplyOutlineNumberDefaultOld()
		{
			 Factory.ExecuteMethod(this, "ApplyOutlineNumberDefaultOld");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		/// <param name="applyTo">optional object applyTo</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ApplyListTemplateOld(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object applyTo)
		{
			 Factory.ExecuteMethod(this, "ApplyListTemplateOld", listTemplate, continuePreviousList, applyTo);
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
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ApplyListTemplateOld(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList)
		{
			 Factory.ExecuteMethod(this, "ApplyListTemplateOld", listTemplate, continuePreviousList);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840622.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ListOutdent()
		{
			 Factory.ExecuteMethod(this, "ListOutdent");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192794.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ListIndent()
		{
			 Factory.ExecuteMethod(this, "ListIndent");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194355.aspx </remarks>
		/// <param name="defaultListBehavior">optional object defaultListBehavior</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ApplyBulletDefault(object defaultListBehavior)
		{
			 Factory.ExecuteMethod(this, "ApplyBulletDefault", defaultListBehavior);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194355.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ApplyBulletDefault()
		{
			 Factory.ExecuteMethod(this, "ApplyBulletDefault");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839320.aspx </remarks>
		/// <param name="defaultListBehavior">optional object defaultListBehavior</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ApplyNumberDefault(object defaultListBehavior)
		{
			 Factory.ExecuteMethod(this, "ApplyNumberDefault", defaultListBehavior);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839320.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ApplyNumberDefault()
		{
			 Factory.ExecuteMethod(this, "ApplyNumberDefault");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822954.aspx </remarks>
		/// <param name="defaultListBehavior">optional object defaultListBehavior</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ApplyOutlineNumberDefault(object defaultListBehavior)
		{
			 Factory.ExecuteMethod(this, "ApplyOutlineNumberDefault", defaultListBehavior);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822954.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ApplyOutlineNumberDefault()
		{
			 Factory.ExecuteMethod(this, "ApplyOutlineNumberDefault");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197184.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		/// <param name="applyTo">optional object applyTo</param>
		/// <param name="defaultListBehavior">optional object defaultListBehavior</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ApplyListTemplate(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object applyTo, object defaultListBehavior)
		{
			 Factory.ExecuteMethod(this, "ApplyListTemplate", listTemplate, continuePreviousList, applyTo, defaultListBehavior);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197184.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197184.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ApplyListTemplate(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList)
		{
			 Factory.ExecuteMethod(this, "ApplyListTemplate", listTemplate, continuePreviousList);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197184.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		/// <param name="applyTo">optional object applyTo</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ApplyListTemplate(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object applyTo)
		{
			 Factory.ExecuteMethod(this, "ApplyListTemplate", listTemplate, continuePreviousList, applyTo);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195913.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		/// <param name="applyTo">optional object applyTo</param>
		/// <param name="defaultListBehavior">optional object defaultListBehavior</param>
		/// <param name="applyLevel">optional object applyLevel</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ApplyListTemplateWithLevel(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object applyTo, object defaultListBehavior, object applyLevel)
		{
			 Factory.ExecuteMethod(this, "ApplyListTemplateWithLevel", new object[]{ listTemplate, continuePreviousList, applyTo, defaultListBehavior, applyLevel });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195913.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195913.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195913.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		/// <param name="applyTo">optional object applyTo</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ApplyListTemplateWithLevel(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object applyTo)
		{
			 Factory.ExecuteMethod(this, "ApplyListTemplateWithLevel", listTemplate, continuePreviousList, applyTo);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195913.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		/// <param name="applyTo">optional object applyTo</param>
		/// <param name="defaultListBehavior">optional object defaultListBehavior</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ApplyListTemplateWithLevel(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object applyTo, object defaultListBehavior)
		{
			 Factory.ExecuteMethod(this, "ApplyListTemplateWithLevel", listTemplate, continuePreviousList, applyTo, defaultListBehavior);
		}

		#endregion

		#pragma warning restore
	}
}
