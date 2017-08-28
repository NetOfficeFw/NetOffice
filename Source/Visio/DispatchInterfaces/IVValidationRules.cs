using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// DispatchInterface IVValidationRules 
	/// SupportByVersion Visio, 14,15,16
	/// </summary>
	[SupportByVersion("Visio", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class IVValidationRules : COMObject , IEnumerable<NetOffice.VisioApi.IVValidationRule>
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
                    _type = typeof(IVValidationRules);
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IVValidationRules(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVValidationRules(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVValidationRules(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVValidationRules(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVValidationRules(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVValidationRules() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVValidationRules(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application", NetOffice.VisioApi.IVApplication.LateBindingApiWrapperType);
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
		public NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "Document", NetOffice.VisioApi.IVDocument.LateBindingApiWrapperType);
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
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="nameUOrIndex">object nameUOrIndex</param>
		[SupportByVersion("Visio", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.VisioApi.IVValidationRule this[object nameUOrIndex]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVValidationRule>(this, "Item", NetOffice.VisioApi.IVValidationRule.LateBindingApiWrapperType, nameUOrIndex);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="ruleID">Int32 ruleID</param>
		[SupportByVersion("Visio", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVValidationRule get_ItemFromID(Int32 ruleID)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVValidationRule>(this, "ItemFromID", NetOffice.VisioApi.IVValidationRule.LateBindingApiWrapperType, ruleID);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Alias for get_ItemFromID
		/// </summary>
		/// <param name="ruleID">Int32 ruleID</param>
		[SupportByVersion("Visio", 14,15,16), Redirect("get_ItemFromID")]
		public NetOffice.VisioApi.IVValidationRule ItemFromID(Int32 ruleID)
		{
			return get_ItemFromID(ruleID);
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="nameU">string nameU</param>
		[SupportByVersion("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVValidationRule Add(string nameU)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.VisioApi.IVValidationRule>(this, "Add", NetOffice.VisioApi.IVValidationRule.LateBindingApiWrapperType, nameU);
		}

		#endregion

       #region IEnumerable<NetOffice.VisioApi.IVValidationRule> Member
        
        /// <summary>
		/// SupportByVersion Visio, 14,15,16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
       public IEnumerator<NetOffice.VisioApi.IVValidationRule> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.VisioApi.IVValidationRule item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersion Visio, 14,15,16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}



