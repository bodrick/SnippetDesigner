using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace Microsoft.SnippetDesigner
{
    public enum CollectionOperation
    {
        Insert,
        Remove,
        Set
    }

    public class CollectionEventArgs<U> : EventArgs
    {
        public CollectionEventArgs(U item, CollectionOperation operation, int index)
        {
            Item = item;
            Index = index;
            Operation = operation;
        }

        public int Index { get; private set; }
        public U Item { get; private set; }
        public CollectionOperation Operation { get; private set; }
    }

    public class CollectionWithEvents<T> : Collection<T>
    {
        public CollectionWithEvents()
        {
        }

        public CollectionWithEvents(IEnumerable<T> items)
            : base(items.ToList())
        {
        }

        public event EventHandler<CollectionEventArgs<T>> CollectionChanged;

        public void AddRange(IEnumerable<T> items)
        {
            foreach (var item in items)
            {
                Add(item);
            }
        }

        protected override void InsertItem(int index, T item)
        {
            base.InsertItem(index, item);
            OnCollectionChanged(item, CollectionOperation.Insert, index);
        }

        protected override void RemoveItem(int index)
        {
            base.RemoveItem(index);
            OnCollectionChanged(default, CollectionOperation.Remove, index);
        }

        protected override void SetItem(int index, T item)
        {
            base.SetItem(index, item);
            OnCollectionChanged(item, CollectionOperation.Set, index);
        }

        private void OnCollectionChanged(T item, CollectionOperation operation, int index) => CollectionChanged?.Invoke(this, new CollectionEventArgs<T>(item, operation, index));
    }
}
