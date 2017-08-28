using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSComctlLibApi
{
	/// <summary>
	/// DispatchInterface IListSubItems 
	/// SupportByVersion MSComctlLib, 6
	/// </summary>
	[SupportByVersion("MSComctlLib", 6)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class IListSubItems : COMObject , IEnumerable<NetOffice.MSComctlLibApi.IListSubItem>
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
                    _type = typeof(IListSubItems);
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IListSubItems(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListSubItems(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListSubItems(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListSubItems(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListSubItems(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListSubItems() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListSubItems(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Count", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSComctlLibApi.IListSubItem get_ControlDefault(object index)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSComctlLibApi.IListSubItem>(this, "ControlDefault", NetOffice.MSComctlLibApi.IListSubItem.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Alias for get_ControlDefault
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6), Redirect("get_ControlDefault")]
		public NetOffice.MSComctlLibApi.IListSubItem ControlDefault(object index)
		{
			return get_ControlDefault(index);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.MSComctlLibApi.IListSubItem this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSComctlLibApi.IListSubItem>(this, "Item", NetOffice.MSComctlLibApi.IListSubItem.LateBindingApiWrapperType, index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		/// <param name="reportIcon">optional object reportIcon</param>
		/// <param name="toolTipText">optional object toolTipText</param>
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IListSubItem Add(object index, object key, object text, object reportIcon, object toolTipText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSComctlLibApi.IListSubItem>(this, "Add", NetOffice.MSComctlLibApi.IListSubItem.LateBindingApiWrapperType, new object[]{ index, key, text, reportIcon, toolTipText });
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IListSubItem Add()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSComctlLibApi.IListSubItem>(this, "Add", NetOffice.MSComctlLibApi.IListSubItem.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IListSubItem Add(object index)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSComctlLibApi.IListSubItem>(this, "Add", NetOffice.MSComctlLibApi.IListSubItem.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IListSubItem Add(object index, object key)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSComctlLibApi.IListSubItem>(this, "Add", NetOffice.MSComctlLibApi.IListSubItem.LateBindingApiWrapperType, index, key);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IListSubItem Add(object index, object key, object text)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSComctlLibApi.IListSubItem>(this, "Add", NetOffice.MSComctlLibApi.IListSubItem.LateBindingApiWrapperType, index, key, text);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		/// <param name="reportIcon">optional object reportIcon</param>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IListSubItem Add(object index, object key, object text, object reportIcon)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSComctlLibApi.IListSubItem>(this, "Add", NetOffice.MSComctlLibApi.IListSubItem.LateBindingApiWrapperType, index, key, text, reportIcon);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public void Clear()
		{
			 Factory.ExecuteMethod(this, "Clear");
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6)]
		public void Remove(object index)
		{
			 Factory.ExecuteMethod(this, "Remove", index);
		}

		#endregion

       #region IEnumerable<NetOffice.MSComctlLibApi.IListSubItem> Member
        
        /// <summary>
		/// SupportByVersion MSComctlLib, 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
       public IEnumerator<NetOffice.MSComctlLibApi.IListSubItem> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.MSComctlLibApi.IListSubItem item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersion MSComctlLib, 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this);
		}

		#endregion

		#pragma warning restore
	}
}



