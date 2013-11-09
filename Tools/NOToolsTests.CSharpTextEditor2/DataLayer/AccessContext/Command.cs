using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    /// <summary>
    /// Defines a command with associated undo action
    /// </summary>
    public class Command
    {
        #region Fields

        private List<Closure[]> _toDoList = new List<Closure[]>();

        #endregion

        #region Ctor
        
        /// <summary>
        /// Creates an instance of th class
        /// </summary>
        /// <param name="name">name of the command</param>
        /// <param name="redo">redo closure</param>
        /// <param name="undo">undo closure</param>
        public Command(string name, Closure redo, Closure undo)
        {
            Name = name;
            this.Add(redo, undo);
        }

        #endregion

        #region Properties

        public string Name { get; private set; }

        internal List<Closure[]> ToDoList
        {
            get
            {
                return _toDoList;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Performs Redo Action
        /// </summary>
        public void Redo()
        {
            _toDoList.ForEach(todo => todo[0].Execute());
        }

        /// <summary>
        /// Performs Undo Action
        /// </summary>
        public void Undo()
        {
            for (int i = _toDoList.Count; i-- > 0;)
                _toDoList[i][1].Execute();
        }

        /// <summary>
        /// Add the redo/undo closures
        /// </summary>
        /// <param name="redo">redo closure</param>
        /// <param name="undo">undo closure</param>
        public void Add(Closure redo, Closure undo)
        {
            _toDoList.Add(new Closure[] { redo, undo });
        }

        #endregion
        
        #region Overrides

        /// <summary>
        /// Returns a String.String that represents the instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {
            if (!String.IsNullOrWhiteSpace(Name))
                return Name;
            else
                return base.ToString();
        }

        #endregion
    }
}
