using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.ADODBApi
{
	///<summary>
	/// DispatchInterface Fields 
	/// SupportByVersion ADODB, 2.1,2.5
	///</summary>
	[SupportByVersionAttribute("ADODB", 2.1,2.5)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Fields : Fields20 ,IEnumerable<NetOffice.ADODBApi.Field>
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
                    _type = typeof(Fields);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Fields(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.1,2.5)]
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
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.ADODBApi.Field this[object index]
		{
			get
{			
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Item", paramsArray);
			NetOffice.ADODBApi.Field newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ADODBApi.Field.LateBindingApiWrapperType) as NetOffice.ADODBApi.Field;
			return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// 
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.1,2.5)]
		public void Refresh()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Refresh", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="type">NetOffice.ADODBApi.Enums.DataTypeEnum Type</param>
		/// <param name="definedSize">optional Int32 DefinedSize = 0</param>
		/// <param name="attrib">optional NetOffice.ADODBApi.Enums.FieldAttributeEnum Attrib = -1</param>
		[SupportByVersionAttribute("ADODB", 2.1,2.5)]
		public void Append(string name, NetOffice.ADODBApi.Enums.DataTypeEnum type, object definedSize, object attrib)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, type, definedSize, attrib);
			Invoker.Method(this, "Append", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="type">NetOffice.ADODBApi.Enums.DataTypeEnum Type</param>
		/// <param name="definedSize">optional Int32 DefinedSize = 0</param>
		/// <param name="attrib">optional NetOffice.ADODBApi.Enums.FieldAttributeEnum Attrib = -1</param>
		/// <param name="fieldValue">optional object FieldValue</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Append(string name, NetOffice.ADODBApi.Enums.DataTypeEnum type, object definedSize, object attrib, object fieldValue)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, type, definedSize, attrib, fieldValue);
			Invoker.Method(this, "Append", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="type">NetOffice.ADODBApi.Enums.DataTypeEnum Type</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.1,2.5)]
		public void Append(string name, NetOffice.ADODBApi.Enums.DataTypeEnum type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, type);
			Invoker.Method(this, "Append", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="type">NetOffice.ADODBApi.Enums.DataTypeEnum Type</param>
		/// <param name="definedSize">optional Int32 DefinedSize = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.1,2.5)]
		public void Append(string name, NetOffice.ADODBApi.Enums.DataTypeEnum type, object definedSize)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, type, definedSize);
			Invoker.Method(this, "Append", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// 
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("ADODB", 2.1)]
		public void Delete(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Update()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Update", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="resyncValues">optional NetOffice.ADODBApi.Enums.ResyncEnum ResyncValues = 2</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Resync(object resyncValues)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(resyncValues);
			Invoker.Method(this, "Resync", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Resync()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Resync", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void CancelUpdate()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CancelUpdate", paramsArray);
		}

		#endregion

       #region IEnumerable<NetOffice.ADODBApi.Field> Member
        
        /// <summary>
		/// SupportByVersionAttribute ADODB, 2.1,2.5
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.1,2.5)]
       public IEnumerator<NetOffice.ADODBApi.Field> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.ADODBApi.Field item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute ADODB, 2.1,2.5
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.1,2.5)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this);
		}

		#endregion
		#pragma warning restore
	}
}