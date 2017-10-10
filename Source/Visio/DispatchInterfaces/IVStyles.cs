using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// DispatchInterface IVStyles 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class IVStyles : COMObject, IEnumerableProvider<NetOffice.VisioApi.IVStyle>
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
                    _type = typeof(IVStyles);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IVStyles(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IVStyles(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVStyles(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVStyles(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVStyles(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVStyles(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVStyles() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVStyles(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 ObjectType
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="nameOrIndex">object nameOrIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.VisioApi.IVStyle this[object nameOrIndex]
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVStyle>(this, "Item", nameOrIndex);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 Count16
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Count16");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "Document");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Stat
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Stat");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVEventList EventList
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVEventList>(this, "EventList");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 PersistsEvents
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "PersistsEvents");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="objectID">Int16 objectID</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVStyle get_ItemFromID16(Int16 objectID)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVStyle>(this, "ItemFromID16", NetOffice.VisioApi.IVStyle.LateBindingApiWrapperType, objectID);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ItemFromID16
		/// </summary>
		/// <param name="objectID">Int16 objectID</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ItemFromID16")]
		public NetOffice.VisioApi.IVStyle ItemFromID16(Int16 objectID)
		{
			return get_ItemFromID16(objectID);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="objectID">Int32 objectID</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVStyle get_ItemFromID(Int32 objectID)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVStyle>(this, "ItemFromID", NetOffice.VisioApi.IVStyle.LateBindingApiWrapperType, objectID);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ItemFromID
		/// </summary>
		/// <param name="objectID">Int32 objectID</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ItemFromID")]
		public NetOffice.VisioApi.IVStyle ItemFromID(Int32 objectID)
		{
			return get_ItemFromID(objectID);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="nameOrIndex">object nameOrIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVStyle get_ItemU(object nameOrIndex)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVStyle>(this, "ItemU", NetOffice.VisioApi.IVStyle.LateBindingApiWrapperType, nameOrIndex);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ItemU
		/// </summary>
		/// <param name="nameOrIndex">object nameOrIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ItemU")]
		public NetOffice.VisioApi.IVStyle ItemU(object nameOrIndex)
		{
			return get_ItemU(nameOrIndex);
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="styleName">string styleName</param>
		/// <param name="basedOn">string basedOn</param>
		/// <param name="fIncludesText">Int16 fIncludesText</param>
		/// <param name="fIncludesLine">Int16 fIncludesLine</param>
		/// <param name="fIncludesFill">Int16 fIncludesFill</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVStyle Add(string styleName, string basedOn, Int16 fIncludesText, Int16 fIncludesLine, Int16 fIncludesFill)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVStyle>(this, "Add", new object[]{ styleName, basedOn, fIncludesText, fIncludesLine, fIncludesFill });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="localeSpecificNameArray">String[] localeSpecificNameArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void GetNames(out String[] localeSpecificNameArray)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			localeSpecificNameArray = null;
			object[] paramsArray = Invoker.ValidateParamsArray((object)localeSpecificNameArray);
			Invoker.Method(this, "GetNames", paramsArray, modifiers);
			localeSpecificNameArray = (String[])paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="localeIndependentNameArray">String[] localeIndependentNameArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void GetNamesU(out String[] localeIndependentNameArray)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			localeIndependentNameArray = null;
			object[] paramsArray = Invoker.ValidateParamsArray((object)localeIndependentNameArray);
			Invoker.Method(this, "GetNamesU", paramsArray, modifiers);
			localeIndependentNameArray = (String[])paramsArray[0];
		}

        #endregion

        #region IEnumerableProvider<NetOffice.VisioApi.IVStyle>

        ICOMObject IEnumerableProvider<NetOffice.VisioApi.IVStyle>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.VisioApi.IVStyle>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.VisioApi.IVStyle>

        /// <summary>
        /// SupportByVersion Visio, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.VisioApi.IVStyle> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.VisioApi.IVStyle item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Visio, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Visio", 11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}