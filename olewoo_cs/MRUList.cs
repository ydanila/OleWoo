using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.Win32;

namespace olewoo_cs
{
    class MRUList
    {
        const int MAXDATAITEMS = 9;

        RegistryKey _mrufiles;
        LinkedList<String> _data;
        Dictionary<String, bool> _uqitems;

        public MRUList(String path)
        {
            RegistryKey hkcu = Registry.CurrentUser;
            _mrufiles = GetRegKey(path, hkcu);
            Refresh();
        }

        public String[] Items
        {
            get
            {
                return _data.ToArray();
            }
        }

        public void Clear()
        {
            _data = new LinkedList<string>();
            _uqitems = new Dictionary<string, bool>(StringComparer.CurrentCultureIgnoreCase);
        }

        public void AddItem(String path)
        {
            if (!_uqitems.ContainsKey(path))
            {
                while (_data.Count >= MAXDATAITEMS)
                {
                    String last = _data.Last.Value;
                    _uqitems.Remove(last);
                    _data.RemoveLast();
                }
                _data.AddFirst(path);
                _uqitems[path] = true;
            }
        }

        private void Refresh()
        {
            Clear();
            for (int i = 0; i < 10; ++i)
            {
                Object o = _mrufiles.GetValue("MruFile" + i);
                String s = o as String;
                if (String.IsNullOrEmpty(s)) return;
                _data.AddLast(s);
                _uqitems[s] = true;
            }
        }

        public void Flush()
        {
            int mval = _data.Count;
            String []data = _data.ToArray();
            for (int i = 0; i < 10; ++i)
            {
                String keyname = "MruFile" + i;
                if (i >= mval)
                {
                    _mrufiles.SetValue(keyname, "");
                }
                else
                {
                    _mrufiles.SetValue(keyname, data[i]);
                }
            }
        }

        private RegistryKey GetRegKey(string key, RegistryKey basekey)
        {
            RegistryKey nkey = basekey.OpenSubKey(key, true);
            if (nkey == null) nkey = basekey.CreateSubKey(key);
            return nkey;
        }
    }
}
