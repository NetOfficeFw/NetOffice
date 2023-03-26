using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface Fields20 
	/// SupportByVersion ADODB, 2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Method)]
	public class Fields20 : Fields15, IEnumerableProvider<object>
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
                    _type = typeof(Fields20);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Fields20(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Fields20(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields20(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields20(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields20(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields20(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields20() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields20(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public void Refresh()
		{
			 Factory.ExecuteMethod(this, "Refresh");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="type">NetOffice.ADODBApi.Enums.DataTypeEnum type</param>
		/// <param name="definedSize">optional Int32 DefinedSize = 0</param>
		/// <param name="attrib">optional NetOffice.ADODBApi.Enums.FieldAttributeEnum Attrib = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("ADODB", 2.5)]
		public void _Append(string name, NetOffice.ADODBApi.Enums.DataTypeEnum type, object definedSize, object attrib)
		{
			 Factory.ExecuteMethod(this, "_Append", name, type, definedSize, attrib);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="type">NetOffice.ADODBApi.Enums.DataTypeEnum type</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void _Append(string name, NetOffice.ADODBApi.Enums.DataTypeEnum type)
		{
			 Factory.ExecuteMethod(this, "_Append", name, type);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="type">NetOffice.ADODBApi.Enums.DataTypeEnum type</param>
		/// <param name="definedSize">optional Int32 DefinedSize = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void _Append(string name, NetOffice.ADODBApi.Enums.DataTypeEnum type, object definedSize)
		{
			 Factory.ExecuteMethod(this, "_Append", name, type, definedSize);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("ADODB", 2.5)]
		public void Delete(object index)
		{
			 Factory.ExecuteMethod(this, "Delete", index);
		}

        #endregion

        #region IEnumerableProvider<object>

        ICOMObject IEnumerableProvider<object>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsMethod(parent, this);
        }

        IEnumerable IEnumerableProvider<object>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, true);
        }

        #endregion

        #region IEnumerable<object>

        /// <summary>
        /// SupportByVersion ADODB, 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.5)]
        public IEnumerator<object> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (object item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion ADODB, 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.5)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this);
		}

		#endregion

		#pragma warning restore
	}
}