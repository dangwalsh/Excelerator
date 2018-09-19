using System.Collections;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;

namespace Gensler.Revit.Excelerator.Views
{
    public class ItemsChangeObservableCollection<T> : ObservableCollection<T> where T : INotifyPropertyChanged
    {
        //public ItemsChangeObservableCollection() : base() { } 
        //public ItemsChangeObservableCollection(IEnumerable<T> collection) : base(collection) { } 

        /// <summary>
        /// Routes item events to the proper binding action
        /// </summary>
        /// <param name="e"></param>
        protected override void OnCollectionChanged(NotifyCollectionChangedEventArgs e)
        {
            switch (e.Action)
            {
                case NotifyCollectionChangedAction.Add:
                    RegisterPropertyChanged(e.NewItems);
                    break;
                case NotifyCollectionChangedAction.Remove:
                    UnregisterPropertyChanged(e.OldItems);
                    break;
                case NotifyCollectionChangedAction.Replace:
                    UnregisterPropertyChanged(e.OldItems);
                    RegisterPropertyChanged(e.NewItems);
                    break;
                case NotifyCollectionChangedAction.Move:
                    break;
                case NotifyCollectionChangedAction.Reset:
                    break;
                default:
                    break;
            }
            base.OnCollectionChanged(e);
        }

        /// <summary>
        /// Removes PropertyChanged binding when clearing the collection
        /// </summary>
        protected override void ClearItems()
        {
            UnregisterPropertyChanged(this);
            base.ClearItems();
        }

        /// <summary>
        /// Binds item PropertyChanged event to item_PropertyChanged
        /// </summary>
        /// <param name="items"></param>
        private void RegisterPropertyChanged(IList items)
        {
            foreach (INotifyPropertyChanged item in items)
            {
                if (item != null)
                {
                    item.PropertyChanged += item_PropertyChanged;
                }
            }
        }

        /// <summary>
        /// Unbinds item PropertyChanged event to item_PropertyChanged
        /// </summary>
        /// <param name="items"></param>
        private void UnregisterPropertyChanged(IList items)
        {
            foreach (INotifyPropertyChanged item in items)
            {
                if (item != null)
                {
                    item.PropertyChanged -= item_PropertyChanged;
                }
            }
        }

        /// <summary>
        /// Event handler that raises OnCollectionChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void item_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            base.OnCollectionChanged(new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset));
        }
    }
}
