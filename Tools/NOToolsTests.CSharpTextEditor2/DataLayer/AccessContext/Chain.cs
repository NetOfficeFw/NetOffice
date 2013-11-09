using System;
using System.ComponentModel;
using System.Collections.Generic;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    /// <summary>
    /// Collection for Command instances to provides an undo/redo for an access context 
    /// </summary>
    /// <typeparam name="T">Command</typeparam>
    public class Chain<T> : INotifyPropertyChanged
    {
        #region Embedded Types

        private class Item
        {
            public Item Prev;
            public Item Next;
            public T Value;
        }

        #endregion

        #region Fields

        private Item _current = new Item();
        private PropertyChangedEventArgs _canForwardArgs = new PropertyChangedEventArgs("CanForward");
        private PropertyChangedEventArgs _canBackwardArgs = new PropertyChangedEventArgs("CanBackward");
        
        #endregion

        #region Properties
        
        /// <summary>
        /// Returns info its possible to step forward
        /// </summary>
        public bool CanForward { get { return _current.Next != null; } }
        
        /// <summary>
        /// Returns info its possible to step backward
        /// </summary>
        public bool CanBackward { get { return _current.Prev != null; } }

        /// <summary>
        /// Count of commands in the collection instance
        /// </summary>
        public int Count
        {
            get
            {
                if (null == _current)
                    return 0;
                int counter = 1;

                Item item = _current.Next;
                while (null != item)
                {
                    counter++;
                    item = item.Next;
                }

                item = _current.Prev;
                while (null != item)
                {
                    counter++;
                    item = item.Prev;
                }

                return counter;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Move one step fwd or back in the command chain
        /// </summary>
        /// <param name="forward">move fwd if true, otherwise false</param>
        /// <returns>the now current command or null</returns>
        public T Move(bool forward)
        {
            if (forward)
            {
                if (CanForward)
                {
                    var couldBack = CanBackward;
                    _current = _current.Next;
                    if (!couldBack)
                        PropertyChanged(this, _canBackwardArgs);
                    if (!CanForward)
                        PropertyChanged(this, _canForwardArgs);
                    // Current-Wert nach Ausführung des Schrittes zurückgeben
                    return _current.Value;
                }
            }
            else
            {
                if (CanBackward)
                {
                    var couldForward = CanForward;
                    _current = _current.Prev;
                    if (!couldForward)
                        PropertyChanged(this, _canForwardArgs);
                    if (!CanBackward)
                        PropertyChanged(this, _canBackwardArgs);
                    //beachte: es wird der Current-Wert vor Ausführung des Schrittes zurückgegeben!
                    return _current.Next.Value;
                }
            }
            return default(T);
        }

        /// <summary>
        /// Clears the collection
        /// </summary>
        public void Clear()
        {
            var couldBack = CanBackward;
            var couldForward = CanForward;
            _current = new Item();
            if (couldForward)
                PropertyChanged(this, _canForwardArgs);
            if (couldBack)
                PropertyChanged(this, _canBackwardArgs);
        }

        /// <summary>
        /// Append a new command to the current command.
        /// When the current command has fwd command these commands are not valid anymore
        /// </summary>
        public virtual void Append(T value)
        {
            _current.Next = new Item() { Prev = _current, Value = value };
            Move(true);
        }

        #endregion

        #region INotifyPropertyChanged

        /// <summary>
        /// Occcurs when the property CanBackward or CanForward has changed
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged = delegate { };

        #endregion
    }
}
