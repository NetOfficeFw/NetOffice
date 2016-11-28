using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace NetOffice
{
    /// <summary>
    /// Sink Helper Base Class for an Event Interface Sink helper class
    /// </summary>
    public abstract class SinkHelper : IDisposable
    {
        #region Fields

        private static List<SinkHelper> _pointList = new List<SinkHelper>();
        private ICOMObject _eventClass;
        private IConnectionPoint _connectionPoint;
        private int _connectionCookie;

        #endregion

        #region Construction

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass">target CoClass instance</param>
        public SinkHelper(ICOMObject eventClass)
        {
            if (null == eventClass)
                throw new ArgumentNullException("eventClass");
            _eventClass = eventClass;
        }

        #endregion

        #region Static Methods

        /// <summary>
        /// Try to find connection point by FindConnectionPoint
        /// </summary>
        private static string FindConnectionPoint(ICOMObject comInstance, IConnectionPointContainer connectionPointContainer, ref IConnectionPoint point, params string[] sinkIds)
        {
            try
            {
                for (int i = sinkIds.Length; i > 0; i--)
                {
                    Guid refGuid = new Guid(sinkIds[i - 1]);
                    IConnectionPoint refPoint = null;
                    connectionPointContainer.FindConnectionPoint(ref refGuid, out refPoint);
                    if (null != refPoint)
                    {
                        point = refPoint;
                        return sinkIds[i - 1];
                    }
                }

                return null;
            }
            catch (Exception throwedException)
            {
                comInstance.Console.WriteException(throwedException);
                return null;
            }
        }

        /// <summary>
        /// try to find connection point by EnumConnectionPoints
        /// </summary>
        private static string EnumConnectionPoint(ICOMObject comInstance, IConnectionPointContainer connectionPointContainer, ref IConnectionPoint point, params string[] sinkIds)
        {
            IConnectionPoint[] points = new IConnectionPoint[1];
            IEnumConnectionPoints enumPoints = null;
            try
            {
                connectionPointContainer.EnumConnectionPoints(out enumPoints);
                while (enumPoints.Next(1, points, IntPtr.Zero) == 0) // S_OK = 0 , S_FALSE = 1
                {
                    if (null == points[0])
                        break;

                    Guid interfaceGuid;
                    points[0].GetConnectionInterface(out interfaceGuid);

                    for (int i = sinkIds.Length; i > 0; i--)
                    {
                        string id = interfaceGuid.ToString().Replace("{", "").Replace("}", "");
                        if (true == sinkIds[i - 1].Equals(id, StringComparison.InvariantCultureIgnoreCase))
                        {
                            Marshal.ReleaseComObject(enumPoints);
                            enumPoints = null;
                            point = points[0];
                            return id;
                        }
                    }
                }
                return null;
            }
            catch (Exception throwedException)
            {
                comInstance.Console.WriteException(throwedException);
                return null;
            }
            finally
            {
                if (null != enumPoints)
                    Marshal.ReleaseComObject(enumPoints);
            }
        }

        /// <summary>
        /// Get supported connection point from comProxy
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static string GetConnectionPoint(ICOMObject comInstance, ref IConnectionPoint point, params string[] sinkIds)
        {
            if (null == sinkIds)
                return null;

            IConnectionPointContainer connectionPointContainer = comInstance.UnderlyingObject as IConnectionPointContainer;
            if (null == connectionPointContainer)
            {
                if (comInstance.Settings.EnableEventDebugOutput)
                    comInstance.Console.WriteLine("Unable to cast IConnectionPointContainer.");
                return null;
            }

            if (comInstance.Settings.EnableEventDebugOutput)
                comInstance.Console.WriteLine(comInstance.UnderlyingTypeName + " -> Call FindConnectionPoint");

            string id = FindConnectionPoint(comInstance, connectionPointContainer, ref point, sinkIds);

            if (comInstance.Settings.EnableEventDebugOutput)
                comInstance.Console.WriteLine(comInstance.UnderlyingTypeName + " -> Call FindConnectionPoint passed");

            if (null == id)
            {
                if (comInstance.Settings.EnableEventDebugOutput)
                    comInstance.Console.WriteLine(comInstance.UnderlyingTypeName + " -> Call EnumConnectionPoint");
                id = EnumConnectionPoint(comInstance, connectionPointContainer, ref point, sinkIds);
                if (comInstance.Settings.EnableEventDebugOutput)
                    comInstance.Console.WriteLine(comInstance.UnderlyingTypeName + " -> Call EnumConnectionPoint passed");
            }

            if (null != id)
                return id;
            else
                throw new COMException("Specified instance doesnt implement the target event interface.");
        }

        /// <summary>
        /// Get supported connection point from comProxy in reverse order to GetConnectionPoint
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static string GetConnectionPoint2(ICOMObject comInstance, ref IConnectionPoint point, params string[] sinkIds)
        {
            if (null == sinkIds)
                return null;

            IConnectionPointContainer connectionPointContainer = comInstance.UnderlyingObject as IConnectionPointContainer;
            if (null == connectionPointContainer)
            {
                if (comInstance.Settings.EnableEventDebugOutput)
                    comInstance.Console.WriteLine("Unable to cast IConnectionPointContainer.");
                return null;
            }

            if (comInstance.Settings.EnableEventDebugOutput)
                comInstance.Console.WriteLine(comInstance.UnderlyingTypeName + " -> Call EnumConnectionPoint");

            string id = EnumConnectionPoint(comInstance, connectionPointContainer, ref point, sinkIds);

            if (comInstance.Settings.EnableEventDebugOutput)
                comInstance.Console.WriteLine(comInstance.UnderlyingTypeName + " -> Call EnumConnectionPoint passed");

            if (null == id)
            {
                if (comInstance.Settings.EnableEventDebugOutput)
                    comInstance.Console.WriteLine(comInstance.UnderlyingTypeName + " -> Call FindConnectionPoint");
                id = FindConnectionPoint(comInstance, connectionPointContainer, ref point, sinkIds);
                if (comInstance.Settings.EnableEventDebugOutput)
                    comInstance.Console.WriteLine(comInstance.UnderlyingTypeName + " -> Call FindConnectionPoint passed");
            }

            if (null != id)
                return id;
            else
                throw new COMException("Specified instance doesnt implement the target event interface.");
        }

        /// <summary>
        /// Dispose all active event bridges
        /// </summary>
        public static void DisposeAll()
        {
            foreach (SinkHelper point in _pointList)
                point.RemoveEventBinding(false);
            _pointList.Clear();
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// create event binding
        /// </summary>
        /// <param name="connectPoint">target connection point</param>
        public void SetupEventBinding(IConnectionPoint connectPoint)
        {
            try
            {
                if (true == Settings.Default.EnableEvents)
                {
                    _connectionPoint = connectPoint;
                    _connectionPoint.Advise(this, out _connectionCookie);
                    _pointList.Add(this);
                }
            }
            catch (Exception throwedException)
            {
                _eventClass.Console.WriteException(throwedException);
                throw (throwedException);
            }
        }

        /// <summary>
        /// Release event binding
        /// </summary>
        public void RemoveEventBinding()
        {
            RemoveEventBinding(true);
        }

        /// <summary>
        /// Release event binding
        /// </summary>
        private void RemoveEventBinding(bool removeFromList)
        {
            if (_connectionCookie != 0)
            {
                try
                {
                    _connectionPoint.Unadvise(_connectionCookie);
                    Marshal.ReleaseComObject(_connectionPoint);
                }
                catch (System.Runtime.InteropServices.COMException throwedException)
                {
                    _eventClass.Console.WriteException(throwedException);
                    ; // RPC server is disconnected or dead
                }
                catch (Exception throwedException)
                {
                    _eventClass.Console.WriteException(throwedException);
                    throw new COMException("An error occured.", throwedException);
                }

                _connectionPoint = null;
                _connectionCookie = 0;

                if (removeFromList)
                    _pointList.Remove(this);
            }
        }

        #endregion

        #region IDisposable Members

        /// <summary>
        /// Remove event binding
        /// </summary>
        public void Dispose()
        {
            RemoveEventBinding();
        }

        #endregion
    }
}