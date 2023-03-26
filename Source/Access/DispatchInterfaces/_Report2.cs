using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _Report2 
	/// SupportByVersion Access, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Method)]
	public class _Report2 : _Report, IEnumerableProvider<object>
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
                    _type = typeof(_Report2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Report2(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Report2(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Report2(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Report2(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Report2(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Report2(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Report2() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Report2(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties
        
		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool AutoResize
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoResize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoResize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool AutoCenter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoCenter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoCenter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool PopUp
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PopUp");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PopUp", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool Modal
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Modal");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Modal", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public byte BorderStyle
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "BorderStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BorderStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool ControlBox
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ControlBox");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ControlBox", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public byte MinMaxButtons
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "MinMaxButtons");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MinMaxButtons", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool CloseButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CloseButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CloseButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public Int16 WindowWidth
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "WindowWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WindowWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public Int16 WindowHeight
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "WindowHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WindowHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public Int16 WindowTop
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "WindowTop");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public Int16 WindowLeft
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "WindowLeft");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public object OpenArgs
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "OpenArgs");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "OpenArgs", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.AccessApi._Printer Printer
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.AccessApi._Printer>(this, "Printer");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Printer", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool Moveable
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Moveable");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Moveable", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool UseDefaultPrinter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseDefaultPrinter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseDefaultPrinter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16), ProxyResult]
		public object Recordset
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Recordset");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Recordset", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string RecordSourceQualifier
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RecordSourceQualifier");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecordSourceQualifier", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string Shape
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Shape");
			}
		}

		#endregion

		#region Methods
        
		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public void Move(object left, object top, object width, object height)
		{
			 Factory.ExecuteMethod(this, "Move", left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public void Move(object left)
		{
			 Factory.ExecuteMethod(this, "Move", left);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public void Move(object left, object top)
		{
			 Factory.ExecuteMethod(this, "Move", left, top);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public void Move(object left, object top, object width)
		{
			 Factory.ExecuteMethod(this, "Move", left, top, width);
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
        /// SupportByVersion Access, 10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public IEnumerator<object> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (object item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Access, 10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Access", 10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this);
		}

		#endregion

		#pragma warning restore
	}
}