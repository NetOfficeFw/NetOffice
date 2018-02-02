using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.ComponentModel;
using NetOffice.IO;
using NetOffice.Exceptions;

namespace NetOffice.Tools.Native.Bridge
{
    /// <summary>
    /// Represents a handle to an unmanaged library.
    /// CdeclHandle does not implement any thread-safe operations.
    /// </summary>
    [DebuggerDisplay("{Name}")]
    public class CdeclHandle : IDisposable
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="underlying">underlying module handle</param>
        /// <param name="folder">folder that contains the library</param>
        /// <param name="name">name of the library</param>
        /// <exception cref="ArgumentOutOfRangeException">underlying is empty</exception>
        /// <exception cref="ArgumentNullException">name is null or empty</exception>
        public CdeclHandle(IntPtr underlying, string folder, string name)
        {
            if (underlying == IntPtr.Zero)
                throw new ArgumentOutOfRangeException("underlying", "Underlying module handle can not be empty.");
            if(String.IsNullOrWhiteSpace(name))
                throw new ArgumentNullException("name", "Name can not be null or empty.");

            Underlying = underlying;
            Folder = folder;
            Name = name;
            Functions = new Dictionary<string, Delegate>();
        }

        /// <summary>
        /// Underyling Module Handle is empty
        /// </summary>
        public bool HandleIsZero
        {
            get
            {
                return Underlying == IntPtr.Zero;
            }
        }

        /// <summary>
        /// Name of the Library
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Folder that contains the library
        /// </summary>
        public string Folder { get; private set; }

        /// <summary>
        /// Underlying Module Handle
        /// </summary>
        private IntPtr Underlying { get; set; }

        /// <summary>
        /// Delegate Cache
        /// </summary>
        private Dictionary<string, Delegate> Functions { get; set; }

        /// <summary>
        /// Returns a function pointer by name. The method is caching the operation.
        /// </summary>
        /// <typeparam name="T">target delegate</typeparam>
        /// <param name="name">name of the method</param>
        /// <returns>delegate to unmanaged method</returns>
        /// <exception cref="Win32Exception">Unable to get proc address or function pointer</exception>
        /// <exception cref="ArgumentNullException">an argument is null or empty</exception>
        /// <exception cref="ObjectDisposedException">instance is already disposed</exception>
        public T GetDelegateForFunctionPointer<T>(string name) where T : class // <= no way for a delegate constraint here
        {
            return GetDelegateForFunctionPointer(name, typeof(T)) as T;
        }

        /// <summary>
        /// Returns a function pointer by name. The method is caching the operation.
        /// </summary>
        /// <param name="name">name of the method</param>
        /// <param name="type">target delegate type</param>
        /// <returns>delegate to unmanaged method</returns>
        /// <exception cref="Win32Exception">Unable to get proc address or function pointer</exception>
        /// <exception cref="ArgumentNullException">an argument is null or empty</exception>
        /// <exception cref="ObjectDisposedException">instance is already disposed</exception>
        public Delegate GetDelegateForFunctionPointer(string name, Type type)
        {
            if (String.IsNullOrWhiteSpace(name))
                throw new ArgumentNullException("name");
            if (null == type)
                throw new ArgumentNullException("type");
            if (HandleIsZero)
                throw new ObjectDisposedException(String.Format("CdeclHandle <{0}>", Name));

            Delegate result = null;
            if (!Functions.ContainsKey(name))
            {
                IntPtr ptr = Interop.GetProcAddress(Underlying, name);
                if (ptr == IntPtr.Zero)
                    throw new Win32Exception(String.Format("Unable to get proc address <{0}> in <{1}>.", name, Name));
                result = Marshal.GetDelegateForFunctionPointer(ptr, type) as Delegate;
                if (null == result)
                    throw new Win32Exception(String.Format("Unable to get function pointer <{0}> in <{1}>.", name, Name));              
                Functions.Add(name, result);
            }
            else
                result = Functions[name];

            return result;
        }

        /// <summary>
        /// Loads an unmanaged library from filesystem
        /// </summary>
        /// <param name="fullFileName">full qualified name of the library file</param>
        /// <param name="fileVersion">optional file version to check major and minor</param>
        /// <returns>handle to library</returns>
        /// <exception cref="FileNotFoundException">File is missing</exception>
        /// <exception cref="Win32Exception">Unable to load library</exception>
        /// <exception cref="FileLoadException">A version mismatch occurs</exception>
        /// <exception cref="ArgumentNullException">fullFileName is null or empty</exception>
        /// <exception cref="NetOfficeIOException">I/O related error</exception>
        public static CdeclHandle LoadLibrary(string fullFileName, Version fileVersion = null)
        {
            if (String.IsNullOrWhiteSpace(fullFileName))
                throw new ArgumentNullException("fullFileName");
            if (!File.Exists(fullFileName))
                throw new FileNotFoundException("File is missing.", fullFileName);

            string folder = IOPath.GetDirectoryName(fullFileName);
            string fileName = IOPath.GetFileName(fullFileName);

            if (null != fileVersion)
            {
                FileVersionInfo version = FileVersionInfo.GetVersionInfo(fullFileName);
                if (version.FileMajorPart != fileVersion.Major ||
                    version.FileMinorPart != fileVersion.Minor)
                {
                    throw new FileLoadException(
                        String.Format("Unable to load library <{0}> because a version mismatch occurs.", fileName));
                }
            }

            IntPtr ptr = Interop.LoadLibrary(fullFileName);
            if (ptr == IntPtr.Zero)
                throw new Win32Exception(String.Format("Unable to load library <{0}>.", fileName));
            
            return new CdeclHandle(ptr, folder, fileName);
        }

        /// <summary>
        /// Loads an unmanaged library from filesystem
        /// </summary>
        /// <typeparam name="T">codebase type</typeparam>
        /// <param name="fileName">name(incl. extension) without path of the library</param>
        /// <param name="fileVersion">optional file version to check major and minor</param>
        /// <returns>handle to library</returns>
        /// <exception cref="FileNotFoundException">File is missing</exception>
        /// <exception cref="Win32Exception">Unable to load library</exception>
        /// <exception cref="FileLoadException">A version mismatch occurs</exception>
        /// <exception cref="ArgumentNullException">a non-optional argument is null or empty</exception>
        /// <exception cref="NetOfficeIOException">I/O related error</exception>
        public static CdeclHandle LoadLibrary<T>(string fileName, Version fileVersion = null)
        {
            return LoadLibrary(typeof(T), fileName, fileVersion);
        }

        /// <summary>
        /// Loads an unmanaged library from filesystem
        /// </summary>
        /// <param name="codebaseType">type to analyze directory/codebase from</param>
        /// <param name="fileName">name(incl. extension) without path of the library</param>
        /// <param name="fileVersion">optional file version to check major and minor</param>
        /// <returns>handle to library</returns>
        /// <exception cref="FileNotFoundException">File is missing</exception>
        /// <exception cref="Win32Exception">Unable to load library</exception>
        /// <exception cref="FileLoadException">A version mismatch occurs</exception>
        /// <exception cref="ArgumentNullException">a non-optional argument is null or empty</exception>
        /// <exception cref="NetOfficeIOException">I/O related error</exception>
        public static CdeclHandle LoadLibrary(Type codebaseType, string fileName, Version fileVersion = null)
        {
            if (null == codebaseType)
                throw new ArgumentNullException("codebaseType");
            if (String.IsNullOrWhiteSpace(fileName))
                throw new ArgumentNullException("fileName");

            string location = codebaseType.Assembly.Location;
            string folderPath = Path.GetDirectoryName(location);
            string fullFileName = Path.Combine(folderPath, fileName);
            return LoadLibrary(fullFileName, fileVersion);
        }

        /// <summary>
        /// Clear Function Pointer Delegate Cache
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public void ClearCache()
        {
            Functions.Clear();
        }

        /// <summary>
        /// Lookup inside delegate cache to determine the given function is cached by the instance
        /// </summary>
        /// <param name="function">given function as any</param>
        /// <returns>true if function is cached and instance is not disposed, otherwise false</returns>
        /// <exception cref="ArgumentNullException">function is null</exception>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public bool IsCachedFunction(Delegate function)
        {
            if (null == function)
                throw new ArgumentNullException("function");

            if (HandleIsZero)
                return false;

            return Functions.ContainsValue(function);
        }

        /// <summary>
        /// Returns the underlying module handle
        /// </summary>
        /// <returns>native win32 module handle</returns>
        [EditorBrowsable(EditorBrowsableState.Never)]
        public IntPtr GetUnderlyingHandle()
        {
            return Underlying;
        }

        /// <summary>
        /// Free the library and clears delegate cache
        /// </summary>
        /// <exception cref="Win32Exception">Unable to free library</exception>
        public void Dispose()
        {
            if (Underlying != IntPtr.Zero)
            {
                if (!Interop.FreeLibrary(Underlying))
                    throw new Win32Exception(String.Format("Unable to free library <{0}>.", Name));
                Underlying = IntPtr.Zero;
                Functions.Clear();
            }
        }

        /// <summary>
        /// Returns a System.String that represents the instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {
            return String.Format("{0}<{1}>", Name, Underlying);
        }
    }
}