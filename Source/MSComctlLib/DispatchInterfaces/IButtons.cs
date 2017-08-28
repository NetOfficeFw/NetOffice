using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSComctlLibApi
{
	/// <summary>
	/// DispatchInterface IButtons 
	/// SupportByVersion MSComctlLib, 6
	/// </summary>
	[SupportByVersion("MSComctlLib", 6)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class IButtons : COMObject , IEnumerable<NetOffice.MSComctlLibApi.IButton>
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
                    _type = typeof(IButtons);
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IButtons(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IButtons(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IButtons(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IButtons(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IButtons(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IButtons() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IButtons(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public Int16 Count
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Count");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Count", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSComctlLibApi.IButton get_ControlDefault(object index)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSComctlLibApi.IButton>(this, "ControlDefault", NetOffice.MSComctlLibApi.IButton.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_ControlDefault(object index, NetOffice.MSComctlLibApi.IButton value)
		{
			Factory.ExecutePropertySet(this, "ControlDefault", index, value);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Alias for get_ControlDefault
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6), Redirect("get_ControlDefault")]
		public NetOffice.MSComctlLibApi.IButton ControlDefault(object index)
		{
			return get_ControlDefault(index);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.MSComctlLibApi.IButton this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSComctlLibApi.IButton>(this, "Item", NetOffice.MSComctlLibApi.IButton.LateBindingApiWrapperType, index);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Item", value, index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6)]
		public void Remove(object index)
		{
			 Factory.ExecuteMethod(this, "Remove", index);
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
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="style">optional object style</param>
		/// <param name="image">optional object image</param>
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IButton Add(object index, object key, object caption, object style, object image)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSComctlLibApi.IButton>(this, "Add", NetOffice.MSComctlLibApi.IButton.LateBindingApiWrapperType, new object[]{ index, key, caption, style, image });
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IButton Add()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSComctlLibApi.IButton>(this, "Add", NetOffice.MSComctlLibApi.IButton.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IButton Add(object index)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSComctlLibApi.IButton>(this, "Add", NetOffice.MSComctlLibApi.IButton.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IButton Add(object index, object key)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSComctlLibApi.IButton>(this, "Add", NetOffice.MSComctlLibApi.IButton.LateBindingApiWrapperType, index, key);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="caption">optional object caption</param>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IButton Add(object index, object key, object caption)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSComctlLibApi.IButton>(this, "Add", NetOffice.MSComctlLibApi.IButton.LateBindingApiWrapperType, index, key, caption);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="style">optional object style</param>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IButton Add(object index, object key, object caption, object style)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSComctlLibApi.IButton>(this, "Add", NetOffice.MSComctlLibApi.IButton.LateBindingApiWrapperType, index, key, caption, style);
		}

		#endregion

       #region IEnumerable<NetOffice.MSComctlLibApi.IButton> Member
        
        /// <summary>
		/// SupportByVersion MSComctlLib, 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
       public IEnumerator<NetOffice.MSComctlLibApi.IButton> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.MSComctlLibApi.IButton item in innerEnumerator)
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



