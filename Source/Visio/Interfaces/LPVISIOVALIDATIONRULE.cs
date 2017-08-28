using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// Interface LPVISIOVALIDATIONRULE 
	/// SupportByVersion Visio, 14,15,16
	/// </summary>
	[SupportByVersion("Visio", 14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class LPVISIOVALIDATIONRULE : COMObject
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
                    _type = typeof(LPVISIOVALIDATIONRULE);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public LPVISIOVALIDATIONRULE(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public LPVISIOVALIDATIONRULE(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOVALIDATIONRULE(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOVALIDATIONRULE(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOVALIDATIONRULE(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOVALIDATIONRULE(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOVALIDATIONRULE() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOVALIDATIONRULE(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public Int16 Stat
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Stat");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "Document");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public Int16 ObjectType
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public Int32 ID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public string NameU
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NameU");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NameU", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public string Category
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Category");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Category", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public string Description
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Description");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Description", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public bool Ignored
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Ignored");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Ignored", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public string FilterExpression
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FilterExpression");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FilterExpression", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public NetOffice.VisioApi.Enums.VisRuleTargets TargetType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRuleTargets>(this, "TargetType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TargetType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public string TestExpression
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TestExpression");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TestExpression", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVValidationRuleSet RuleSet
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVValidationRuleSet>(this, "RuleSet");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="targetPage">optional NetOffice.VisioApi.IVPage targetPage</param>
		/// <param name="targetShape">optional NetOffice.VisioApi.IVShape targetShape</param>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVValidationIssue AddIssue(object targetPage, object targetShape)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVValidationIssue>(this, "AddIssue", targetPage, targetShape);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVValidationIssue AddIssue()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVValidationIssue>(this, "AddIssue");
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="targetPage">optional NetOffice.VisioApi.IVPage targetPage</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVValidationIssue AddIssue(object targetPage)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVValidationIssue>(this, "AddIssue", targetPage);
		}

		#endregion

		#pragma warning restore
	}
}
