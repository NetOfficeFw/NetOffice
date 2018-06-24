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
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "MSProject", 11, 12, 14), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("BF6D2103-92D3-4162-9816-A3D811BCF8CA")]
	public interface TableFields : ICOMObject, IEnumerableProvider<NetOffice.MSProjectApi.TableField>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Project Parent { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSProject", 11,12,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.MSProjectApi.TableField this[object index] { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Int32 Count { get; }

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
		NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData, object width, object title, object alignTitle, object before, object autoWrap);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		/// <param name="alignData">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignData = 2</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		/// <param name="alignData">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignData = 2</param>
		/// <param name="width">optional Int32 Width = 10</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData, object width);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		/// <param name="alignData">optional NetOffice.MSProjectApi.Enums.PjAlignment AlignData = 2</param>
		/// <param name="width">optional Int32 Width = 10</param>
		/// <param name="title">optional string Title = </param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData, object width, object title);

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
		NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData, object width, object title, object alignTitle);

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
		NetOffice.MSProjectApi.TableField Add(NetOffice.MSProjectApi.Enums.PjField field, object alignData, object width, object title, object alignTitle, object before);

        #endregion

        #region IEnumerable<NetOffice.MSProjectApi.TableField>

        /// <summary>
        /// SupportByVersion MSProject, 11,12,14
        /// </summary>
        [SupportByVersion("MSProject", 11, 12, 14)]
        new IEnumerator<NetOffice.MSProjectApi.TableField> GetEnumerator();

        #endregion
    }
}
