using System.Collections.Generic;
using System.ComponentModel;
using System.Dynamic;

namespace ExcelImportExport
{
    public class DynamicDictionaryWrapper : DynamicObject, INotifyPropertyChanged
    {
        protected readonly Dictionary<string, object> dictionary;

        public DynamicDictionaryWrapper(Dictionary<string, object> dictionary)
        {
            this.dictionary = dictionary;
        }

        public object this[string key]
        {
            get
            {
                if (dictionary.ContainsKey(key))
                {
                    return dictionary[key];
                }
                return null;
            }
            set
            {
                dictionary[key] = value;
                OnPropertyChanged(key);
            }
        }

        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            result = null;
            return (dictionary.TryGetValue(binder.Name, out result));
        }

        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            dictionary[binder.Name] = value;
            OnPropertyChanged(binder.Name);
            return true;
        }

        public override IEnumerable<string> GetDynamicMemberNames()
        {
            return dictionary.Keys;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string propertyName)
        {
            var propChange = PropertyChanged;
            if (propChange == null) return;
            propChange(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
