using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface CommandBarControls 
	/// SupportByVersion Office, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862747.aspx </remarks>
	[SupportByVersion("Office", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class CommandBarControls : _IMsoDispObj , IEnumerable<NetOffice.OfficeApi.CommandBarControl>
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
                    _type = typeof(CommandBarControls);
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public CommandBarControls(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControls(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControls(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControls(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControls(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControls() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControls(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860596.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.OfficeApi.CommandBarControl this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBarControl>(this, "Item", NetOffice.OfficeApi.CommandBarControl.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860798.aspx </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar Parent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBar>(this, "Parent", NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861771.aspx </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="id">optional object id</param>
		/// <param name="parameter">optional object parameter</param>
		/// <param name="before">optional object before</param>
		/// <param name="temporary">optional object temporary</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl Add(object type, object id, object parameter, object before, object temporary)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Add", NetOffice.OfficeApi.CommandBarControl.LateBindingApiWrapperType, new object[]{ type, id, parameter, before, temporary });
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861771.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl Add()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Add", NetOffice.OfficeApi.CommandBarControl.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861771.aspx </remarks>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl Add(object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Add", NetOffice.OfficeApi.CommandBarControl.LateBindingApiWrapperType, type);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861771.aspx </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="id">optional object id</param>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl Add(object type, object id)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Add", NetOffice.OfficeApi.CommandBarControl.LateBindingApiWrapperType, type, id);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861771.aspx </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="id">optional object id</param>
		/// <param name="parameter">optional object parameter</param>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl Add(object type, object id, object parameter)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Add", NetOffice.OfficeApi.CommandBarControl.LateBindingApiWrapperType, type, id, parameter);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861771.aspx </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="id">optional object id</param>
		/// <param name="parameter">optional object parameter</param>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl Add(object type, object id, object parameter, object before)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Add", NetOffice.OfficeApi.CommandBarControl.LateBindingApiWrapperType, type, id, parameter, before);
		}

		#endregion

       #region IEnumerable<NetOffice.OfficeApi.CommandBarControl> Member
        
        /// <summary>
		/// SupportByVersion Office, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
       public IEnumerator<NetOffice.OfficeApi.CommandBarControl> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.OfficeApi.CommandBarControl item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersion Office, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}



