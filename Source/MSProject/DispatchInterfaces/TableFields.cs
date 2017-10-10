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
	/// DispatchInterface TableFields 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920711(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class TableFields : COMObject, IEnumerableProvider<NetOffice.MSProjectApi.TableField>
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
                    _type = typeof(TableFields);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public TableFields(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public TableFields(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableFields(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableFields(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableFields(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableFields(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableFields() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableFields(string progId) : base(progId)
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
		public NetOffice.MSProjectApi.Project Parent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Project>(this, "Parent", NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSProject", 11,12,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.MSProjectApi.TableField this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.TableField>(this, "Item", NetOffice.MSProjectApi.TableField.LateBindingApiWrapperType, index);
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		/// <param name="alignData">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignData = 2</param>
		/// <param name="width">optional Int32 Width = 10</param>
		/// <param name="title">optional string Title = </param>
		/// <param name="alignTitle">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignTitle = 1</param>
		/// <param name="before">optional Int32 Before = -1</param>
		/// <param name="autoWrap">optional bool AutoWrap = true</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData, object width, object title, object alignTitle, object before, object autoWrap)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TableField>(this, "Add", NetOffice.MSProjectApi.TableField.LateBindingApiWrapperType, new object[]{ field, alignData, width, title, alignTitle, before, autoWrap });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TableField>(this, "Add", NetOffice.MSProjectApi.TableField.LateBindingApiWrapperType, field);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		/// <param name="alignData">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignData = 2</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TableField>(this, "Add", NetOffice.MSProjectApi.TableField.LateBindingApiWrapperType, field, alignData);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		/// <param name="alignData">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignData = 2</param>
		/// <param name="width">optional Int32 Width = 10</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TableField>(this, "Add", NetOffice.MSProjectApi.TableField.LateBindingApiWrapperType, field, alignData, width);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		/// <param name="alignData">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignData = 2</param>
		/// <param name="width">optional Int32 Width = 10</param>
		/// <param name="title">optional string Title = </param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData, object width, object title)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TableField>(this, "Add", NetOffice.MSProjectApi.TableField.LateBindingApiWrapperType, field, alignData, width, title);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		/// <param name="alignData">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignData = 2</param>
		/// <param name="width">optional Int32 Width = 10</param>
		/// <param name="title">optional string Title = </param>
		/// <param name="alignTitle">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignTitle = 1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData, object width, object title, object alignTitle)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TableField>(this, "Add", NetOffice.MSProjectApi.TableField.LateBindingApiWrapperType, new object[]{ field, alignData, width, title, alignTitle });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		/// <param name="alignData">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignData = 2</param>
		/// <param name="width">optional Int32 Width = 10</param>
		/// <param name="title">optional string Title = </param>
		/// <param name="alignTitle">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignTitle = 1</param>
		/// <param name="before">optional Int32 Before = -1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData, object width, object title, object alignTitle, object before)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TableField>(this, "Add", NetOffice.MSProjectApi.TableField.LateBindingApiWrapperType, new object[]{ field, alignData, width, title, alignTitle, before });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.MSProjectApi.TableField>

        ICOMObject IEnumerableProvider<NetOffice.MSProjectApi.TableField>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.MSProjectApi.TableField>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.MSProjectApi.TableField>

        /// <summary>
        /// SupportByVersion MSProject, 11,12,14
        /// </summary>
        [SupportByVersion("MSProject", 11, 12, 14)]
        public IEnumerator<NetOffice.MSProjectApi.TableField> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.MSProjectApi.TableField item in innerEnumerator)
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