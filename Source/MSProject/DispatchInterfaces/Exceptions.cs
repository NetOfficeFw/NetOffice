using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.MSProjectApi
{
	/// <summary>
	/// DispatchInterface Exceptions 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920590(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class Exceptions : COMObject, IEnumerableProvider<NetOffice.MSProjectApi.Exception>
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
                    _type = typeof(Exceptions);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Exceptions(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Exceptions(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Exceptions(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Exceptions(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Exceptions(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Exceptions(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Exceptions() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Exceptions(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Application>(this, "Application", NetOffice.MSProjectApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Calendar Parent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Calendar>(this, "Parent", NetOffice.MSProjectApi.Calendar.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSProject", 11,12,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.MSProjectApi.Exception this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Exception>(this, "Item", NetOffice.MSProjectApi.Exception.LateBindingApiWrapperType, index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		/// <param name="name">optional object name</param>
		/// <param name="period">optional object period</param>
		/// <param name="daysOfWeek">optional object daysOfWeek</param>
		/// <param name="monthPosition">optional object monthPosition</param>
		/// <param name="monthItem">optional object monthItem</param>
		/// <param name="month">optional object month</param>
		/// <param name="monthDay">optional object monthDay</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period, object daysOfWeek, object monthPosition, object monthItem, object month, object monthDay)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Exception>(this, "Add", NetOffice.MSProjectApi.Exception.LateBindingApiWrapperType, new object[]{ type, start, finish, occurrences, name, period, daysOfWeek, monthPosition, monthItem, month, monthDay });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Exception>(this, "Add", NetOffice.MSProjectApi.Exception.LateBindingApiWrapperType, type, start);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Exception>(this, "Add", NetOffice.MSProjectApi.Exception.LateBindingApiWrapperType, type, start, finish);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Exception>(this, "Add", NetOffice.MSProjectApi.Exception.LateBindingApiWrapperType, type, start, finish, occurrences);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Exception>(this, "Add", NetOffice.MSProjectApi.Exception.LateBindingApiWrapperType, new object[]{ type, start, finish, occurrences, name });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		/// <param name="name">optional object name</param>
		/// <param name="period">optional object period</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Exception>(this, "Add", NetOffice.MSProjectApi.Exception.LateBindingApiWrapperType, new object[]{ type, start, finish, occurrences, name, period });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		/// <param name="name">optional object name</param>
		/// <param name="period">optional object period</param>
		/// <param name="daysOfWeek">optional object daysOfWeek</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period, object daysOfWeek)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Exception>(this, "Add", NetOffice.MSProjectApi.Exception.LateBindingApiWrapperType, new object[]{ type, start, finish, occurrences, name, period, daysOfWeek });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		/// <param name="name">optional object name</param>
		/// <param name="period">optional object period</param>
		/// <param name="daysOfWeek">optional object daysOfWeek</param>
		/// <param name="monthPosition">optional object monthPosition</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period, object daysOfWeek, object monthPosition)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Exception>(this, "Add", NetOffice.MSProjectApi.Exception.LateBindingApiWrapperType, new object[]{ type, start, finish, occurrences, name, period, daysOfWeek, monthPosition });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		/// <param name="name">optional object name</param>
		/// <param name="period">optional object period</param>
		/// <param name="daysOfWeek">optional object daysOfWeek</param>
		/// <param name="monthPosition">optional object monthPosition</param>
		/// <param name="monthItem">optional object monthItem</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period, object daysOfWeek, object monthPosition, object monthItem)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Exception>(this, "Add", NetOffice.MSProjectApi.Exception.LateBindingApiWrapperType, new object[]{ type, start, finish, occurrences, name, period, daysOfWeek, monthPosition, monthItem });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		/// <param name="name">optional object name</param>
		/// <param name="period">optional object period</param>
		/// <param name="daysOfWeek">optional object daysOfWeek</param>
		/// <param name="monthPosition">optional object monthPosition</param>
		/// <param name="monthItem">optional object monthItem</param>
		/// <param name="month">optional object month</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period, object daysOfWeek, object monthPosition, object monthItem, object month)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Exception>(this, "Add", NetOffice.MSProjectApi.Exception.LateBindingApiWrapperType, new object[]{ type, start, finish, occurrences, name, period, daysOfWeek, monthPosition, monthItem, month });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.MSProjectApi.Exception>

        ICOMObject IEnumerableProvider<NetOffice.MSProjectApi.Exception>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.MSProjectApi.Exception>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.MSProjectApi.Exception>

        /// <summary>
        /// SupportByVersion MSProject, 11,12,14
        /// </summary>
        [SupportByVersion("MSProject", 11, 12, 14)]
        public IEnumerator<NetOffice.MSProjectApi.Exception> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.MSProjectApi.Exception item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion MSProject, 11,12,14
        /// </summary>
        [SupportByVersion("MSProject", 11,12,14)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}