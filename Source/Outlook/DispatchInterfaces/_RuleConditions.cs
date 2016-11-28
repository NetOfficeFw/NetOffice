using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.OutlookApi
{
	///<summary>
	/// DispatchInterface _RuleConditions 
	/// SupportByVersion Outlook, 12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Outlook", 12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _RuleConditions : COMObject ,IEnumerable<NetOffice.OutlookApi._RuleCondition>
	{
		#pragma warning disable
		#region Type Information

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(_RuleConditions);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _RuleConditions(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _RuleConditions(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _RuleConditions(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _RuleConditions(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _RuleConditions(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _RuleConditions() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _RuleConditions(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869062.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi._Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.OutlookApi._Application newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi._Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869290.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlObjectClass Class
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Class", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OutlookApi.Enums.OlObjectClass)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860751.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi._NameSpace Session
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Session", paramsArray);
				NetOffice.OutlookApi._NameSpace newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi._NameSpace;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861035.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866889.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public Int32 Count
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860414.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.RuleCondition CC
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CC", paramsArray);
				NetOffice.OutlookApi.RuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.RuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.RuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869301.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.RuleCondition HasAttachment
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasAttachment", paramsArray);
				NetOffice.OutlookApi.RuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.RuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.RuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865805.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.ImportanceRuleCondition Importance
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Importance", paramsArray);
				NetOffice.OutlookApi.ImportanceRuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.ImportanceRuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.ImportanceRuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860351.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.RuleCondition MeetingInviteOrUpdate
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MeetingInviteOrUpdate", paramsArray);
				NetOffice.OutlookApi.RuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.RuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.RuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868020.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.RuleCondition NotTo
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NotTo", paramsArray);
				NetOffice.OutlookApi.RuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.RuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.RuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862190.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.RuleCondition OnlyToMe
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OnlyToMe", paramsArray);
				NetOffice.OutlookApi.RuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.RuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.RuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868912.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.RuleCondition ToMe
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ToMe", paramsArray);
				NetOffice.OutlookApi.RuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.RuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.RuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862688.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.RuleCondition ToOrCc
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ToOrCc", paramsArray);
				NetOffice.OutlookApi.RuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.RuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.RuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868201.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.AccountRuleCondition Account
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Account", paramsArray);
				NetOffice.OutlookApi.AccountRuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.AccountRuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.AccountRuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868757.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.TextRuleCondition Body
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Body", paramsArray);
				NetOffice.OutlookApi.TextRuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.TextRuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.TextRuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869178.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.TextRuleCondition BodyOrSubject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BodyOrSubject", paramsArray);
				NetOffice.OutlookApi.TextRuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.TextRuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.TextRuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869855.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.CategoryRuleCondition Category
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Category", paramsArray);
				NetOffice.OutlookApi.CategoryRuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.CategoryRuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.CategoryRuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868215.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.FormNameRuleCondition FormName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormName", paramsArray);
				NetOffice.OutlookApi.FormNameRuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.FormNameRuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.FormNameRuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863931.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.ToOrFromRuleCondition From
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "From", paramsArray);
				NetOffice.OutlookApi.ToOrFromRuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.ToOrFromRuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.ToOrFromRuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863091.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.TextRuleCondition MessageHeader
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MessageHeader", paramsArray);
				NetOffice.OutlookApi.TextRuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.TextRuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.TextRuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861916.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.AddressRuleCondition RecipientAddress
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RecipientAddress", paramsArray);
				NetOffice.OutlookApi.AddressRuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.AddressRuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.AddressRuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866273.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.AddressRuleCondition SenderAddress
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SenderAddress", paramsArray);
				NetOffice.OutlookApi.AddressRuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.AddressRuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.AddressRuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868863.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.SenderInAddressListRuleCondition SenderInAddressList
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SenderInAddressList", paramsArray);
				NetOffice.OutlookApi.SenderInAddressListRuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.SenderInAddressListRuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.SenderInAddressListRuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869345.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.TextRuleCondition Subject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Subject", paramsArray);
				NetOffice.OutlookApi.TextRuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.TextRuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.TextRuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865065.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.ToOrFromRuleCondition SentTo
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SentTo", paramsArray);
				NetOffice.OutlookApi.ToOrFromRuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.ToOrFromRuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.ToOrFromRuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866593.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.RuleCondition OnLocalMachine
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OnLocalMachine", paramsArray);
				NetOffice.OutlookApi.RuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.RuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.RuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860398.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.RuleCondition OnOtherMachine
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OnOtherMachine", paramsArray);
				NetOffice.OutlookApi.RuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.RuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.RuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868588.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.RuleCondition AnyCategory
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AnyCategory", paramsArray);
				NetOffice.OutlookApi.RuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.RuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.RuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869523.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.RuleCondition FromAnyRSSFeed
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FromAnyRSSFeed", paramsArray);
				NetOffice.OutlookApi.RuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.RuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.RuleCondition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869818.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.FromRssFeedRuleCondition FromRssFeed
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FromRssFeed", paramsArray);
				NetOffice.OutlookApi.FromRssFeedRuleCondition newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.FromRssFeedRuleCondition.LateBindingApiWrapperType) as NetOffice.OutlookApi.FromRssFeedRuleCondition;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.OutlookApi._RuleCondition this[Int32 index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.OutlookApi._RuleCondition newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi._RuleCondition;
				return newObject;
			}
		}

		#endregion
       #region IEnumerable<NetOffice.OutlookApi._RuleCondition> Member
        
        /// <summary>
		/// SupportByVersionAttribute Outlook, 12,14,15,16
		/// This is a custom enumerator from NetOffice
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
        [CustomEnumerator]
       public IEnumerator<NetOffice.OutlookApi._RuleCondition> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.OutlookApi._RuleCondition item in innerEnumerator)
               yield return item;
       }

       #endregion
   
       #region IEnumerable Members
        
       /// <summary>
		/// SupportByVersionAttribute Outlook, 12,14,15,16
		/// This is a custom enumerator from NetOffice
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
        [CustomEnumerator]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
       {
            int count = Count;
            object[] enumeratorObjects = new object[count];
            for (int i = 0; i < count; i++)
                enumeratorObjects[i] = this[i+1];

            foreach (object item in enumeratorObjects)
                yield return item;
       }

       #endregion
       		#pragma warning restore
	}
}