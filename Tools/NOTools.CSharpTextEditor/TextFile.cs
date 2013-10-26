using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ICSharpCode.AvalonEdit.CodeCompletion;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.SharpDevelop.Dom;

namespace NOTools.CSharpTextEditor
{
    public class TextFile
    {
        #region Fields

        private string _scriptContent;
        private ProjectContentRegistry _contentRegistry;
        private IFilterStrategy _filterStrategy;
        private IProjectContent _projectContent;
        private TextDocument _codeDocument;
        private Dictionary<string, IProjectContent> _cacheReferences = new Dictionary<string, IProjectContent>();

        #endregion

        #region Ctor

        public TextFile()
        {
            _scriptContent = string.Empty;
            _codeDocument = new TextDocument(new StringTextSource(_scriptContent));
            _contentRegistry = new ProjectContentRegistry();
            _filterStrategy = new NonFilterStrategy();
            _projectContent = new DefaultProjectContent();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Assembly info cache path
        /// </summary>
        internal string PersistancePath { get; set; }

        /// <summary>
        /// Gets or sets the content of the script.
        /// </summary>
        public string ScriptContent
        {
            get { return _codeDocument.Text; }
        }

        /// <summary>
        /// Gets the current filter strategy for intellisense.
        /// </summary>
        public IFilterStrategy FilterStrategy
        {
            get { return _filterStrategy; }
        }

        /// <summary>
        /// Gets the IProjectContent object describing all code items
        /// that may be accessed through intellisense.
        /// </summary>
        public IProjectContent ProjectContent
        {
            get { return _projectContent; }
        }

        /// <summary>
        /// Gets the TextDocument object containing the source code.
        /// </summary>
        public TextDocument CodeDocument
        {
            get { return _codeDocument; }
        }

        /// <summary>
        /// Returns current codebase folder
        /// </summary>
        /// <returns>Folder path</returns>
        private string CurrentPath
        {
            get
            {
                if (null == _currentPath)
                {
                    string fileName = typeof(TextFile).Assembly.Location;
                    _currentPath = System.IO.Path.GetDirectoryName(fileName);
                }
                return _currentPath;
            }
        }
        private string _currentPath;

        #endregion

        public DomPersistence Persistence 
        {
            get 
            {
                return _contentRegistry.ActivatePersistence(GetPersistencePath());
            }
        }

        #region Methods

        public bool AddReferenceFromPersistenceFolder(string assemblyName, bool doAsync = false)
        {
            if (!Directory.Exists(GetPersistencePath()))
                Directory.CreateDirectory(GetPersistencePath());

            bool result = false;
            RunMethod(
                  delegate
                  {
                      lock (_contentRegistry)
                      {
                          DomPersistence persistence = _contentRegistry.ActivatePersistence(GetPersistencePath());
                          
                          IProjectContent persistenceContent = persistence.LoadProjectContentByAssemblyName(assemblyName);
                          if (null != persistenceContent)
                          {
                              _cacheReferences[assemblyName] = persistenceContent;
                              //_projectContent.AddReferencedContent(persistenceContent);
                              _projectContent.ReferencedContents.Add(persistenceContent);
                              result = true;
                          }
                      }
                  }, doAsync);
            return result;
        }        

        public void AddReferencesFromPersistenceFolder(string[] assemblyNames, bool doAsync = false)
        {
            RunMethod(
               delegate
               {
                   lock (_contentRegistry)
                   {
                       if (!Directory.Exists(GetPersistencePath()))
                           Directory.CreateDirectory(GetPersistencePath());

                       foreach (string item in assemblyNames)
                       {
                           DomPersistence persistence = _contentRegistry.ActivatePersistence(GetPersistencePath());

                           if (_cacheReferences.ContainsKey(item))
                               continue;

                           IProjectContent persistenceContent = persistence.LoadProjectContentByAssemblyName(item);
                           if (null != persistenceContent)
                           {
                               _cacheReferences[item] = persistenceContent;
                               _projectContent.ReferencedContents.Add(persistenceContent);
                               //_projectContent.AddReferencedContent(persistenceContent);
                           }
                       }
                   }
               }, doAsync);
        }
         
        public void AddReferenceFromFile(string assemblyName, string assemblyLocation, bool tryPersistence = true, bool doAsync = false)
        {
            //if (_cacheReferences.ContainsKey(assemblyName))
            //    return;

            if (!Directory.Exists(GetPersistencePath()))
                Directory.CreateDirectory(GetPersistencePath());

            RunMethod(
               delegate
               {
                   lock (_contentRegistry)
                   {
                       DomPersistence persistence = _contentRegistry.ActivatePersistence(GetPersistencePath());

                       if (true == tryPersistence)
                       {
                           IProjectContent persistenceContent = persistence.LoadProjectContentByAssemblyName(assemblyName);
                           if (null != persistenceContent)
                           {
                               _cacheReferences[assemblyName] = persistenceContent;
                               //_projectContent.AddReferencedContent(persistenceContent);
                               _projectContent.ReferencedContents.Add(persistenceContent);
                               return;
                           }
                       }

                       IProjectContent result = _contentRegistry.GetProjectContentForReference(assemblyName, assemblyLocation);
                       if (null != result)
                       {
                           _cacheReferences[assemblyName] = result;
                           _projectContent.ReferencedContents.Add(result);
                           //_projectContent.AddReferencedContent(result);
                       }
                   }
               }, doAsync);
        }

        public void AddReferencesFromFile(string[] assemblyNames, string[] assemblyLocations, bool tryPersistence = true, bool doAsync = false)
        {
            if (!Directory.Exists(GetPersistencePath()))
                Directory.CreateDirectory(GetPersistencePath());

            RunMethod(
               delegate
               {
                   lock (_contentRegistry)
                   {
                       DomPersistence persistence = _contentRegistry.ActivatePersistence(GetPersistencePath());

                       for (int i = 0; i < assemblyNames.Length; i++)
                       {
                           string assemblyName = assemblyNames[i];
                           string assemblyLocation = assemblyLocations[i];

                           if (_cacheReferences.ContainsKey(assemblyName))
                               continue;

                           if (true == tryPersistence)
                           {
                               IProjectContent persistenceContent = persistence.LoadProjectContentByAssemblyName(assemblyName);
                               if (null != persistenceContent)
                               {
                                   _cacheReferences[assemblyName] = persistenceContent;
                                   //_projectContent.AddReferencedContent(persistenceContent);
                                   _projectContent.ReferencedContents.Add(persistenceContent);

                                   continue;
                               }
                           }

                           IProjectContent result = _contentRegistry.GetProjectContentForReference(assemblyName, assemblyLocation);
                           if (null != result)
                           {
                               _cacheReferences[assemblyName] = result;
                                _projectContent.ReferencedContents.Add(result);
                               // _projectContent.AddReferencedContent(result);
                           }
                       }
                   }
               }, doAsync);
        }

        public void AddWellKnownReferenceFromFile(string assemblyName, string assemblyLocation, bool tryPersistence = true, bool doAsync = false)
        {
            if (_cacheReferences.ContainsKey(assemblyName))
                return;

            if (!Directory.Exists(GetPersistencePath()))
                Directory.CreateDirectory(GetPersistencePath());

            RunMethod(
               delegate
               {
                   lock (_contentRegistry)
                   {
                       DomPersistence persistence = _contentRegistry.ActivatePersistence(GetPersistencePath());

                       if (true == tryPersistence)
                       {
                           IProjectContent persistenceContent = persistence.LoadProjectContentByAssemblyName(assemblyName);
                           if (null != persistenceContent)
                           {
                               _cacheReferences[assemblyName] = persistenceContent;
                               //_projectContent.AddReferencedContent(persistenceContent);
                               _projectContent.ReferencedContents.Add(persistenceContent);
                               return;
                           }
                       }

                       IProjectContent result = _contentRegistry.GetProjectContentForReference(assemblyName, assemblyLocation);
                       if (null != result)
                       {
                           _cacheReferences[assemblyName] = result;
                           _projectContent.ReferencedContents.Add(result);
                           //_projectContent.AddReferencedContent(result);
                       }
                   }
               }, doAsync);
        }

        public void RemoveReference(string assemblyName)
        {
            IProjectContent prjContent = null;
            foreach (IProjectContent item in _projectContent.ReferencedContents)
	        {
                if (item.AssemblyName.Equals(assemblyName, StringComparison.InvariantCultureIgnoreCase))
                {
                    prjContent = item;
                    break;
                }
	        }

            if (null != prjContent)
                _projectContent.ReferencedContents.Remove(prjContent);


        }

        private string CalculatePathOneUpward(string path)
        {
            if(path.EndsWith("\\"))
                path = path.Substring(0, path.Length -1);
            int pos = path.LastIndexOf("\\");
            if (pos > -1)
                path = path.Substring(0, pos);
            return path;
        }

        private string GetPersistencePath()
        {
            if (String.IsNullOrWhiteSpace(PersistancePath))
                return CurrentPath;
            else
            {
                if (PersistancePath.Contains(".."))
                {
                    string relativePath = CurrentPath;
                    string[] array = PersistancePath.Split(new string[] { "\\" }, StringSplitOptions.None);
                    foreach (var item in array)
                    {
                        if (item.Trim() == "..")
                        {
                            relativePath = CalculatePathOneUpward(relativePath); 
                        }
                        else
                        {
                            relativePath = Path.Combine(relativePath, item);
                        }
                    }
                    return relativePath;
                }
                else
                    return PersistancePath;
            }
        }

        private void RunMethod(ThreadStart method, bool runAsync)
        {
            if (runAsync)
            {
                Thread thread1 = new Thread(method);
                thread1.Start();
            }
            else
            {
                method();
            }
        }

        #endregion
    }
}
