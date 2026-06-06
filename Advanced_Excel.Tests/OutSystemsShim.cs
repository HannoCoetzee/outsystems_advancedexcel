// -----------------------------------------------------------------------
// OutSystems.RuntimeCommon, OutSystems.HubEdition.RuntimePlatform shims
// Minimal stubs so that Advanced_Excel source files compile without the
// actual OutSystems runtime DLLs. These provide ONLY the types and method
// signatures referenced by the production code — no real implementation.
// -----------------------------------------------------------------------
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Reflection;
using System.Xml;
using System.Xml.Serialization;

// ---------- OutSystems.ObjectKeys ----------
namespace OutSystems.ObjectKeys
{
    public class GlobalObjectKey
    {
        private readonly string _key;
        private GlobalObjectKey(string key) { _key = key; }
        public static GlobalObjectKey Parse(string key) => new GlobalObjectKey(key);
        public static implicit operator string(GlobalObjectKey k) => k?._key;
        public override string ToString() => _key;
    }

    [AttributeUsage(AttributeTargets.All)]
    public class ObjectKeyAttribute : Attribute
    {
        public ObjectKeyAttribute(string key) { }
    }
}

// ---------- OutSystems.RuntimeCommon ----------
namespace OutSystems.RuntimeCommon
{
    public interface ITypedRecord<T> where T : struct
    {
        void Serialize(object info, SerializationInfo si);
    }

    public static class Information
    {
        public static string GetAttribute(string name) => "";
        public static bool IsDefined(Type t, string name) => false;
    }

    public enum ETypedEvent
    {
        None = 0,
        OnChange = 1,
    }
}

// ---------- OutSystems.HubEdition.RuntimePlatform ----------
namespace OutSystems.HubEdition.RuntimePlatform
{
    public interface IRecord
    {
    } // minimal — used only in signatures

    public class RecordList : ArrayList, IList<IRecord>
    {
        public void Reset() { /* noop */ }
        public void Append(IRecord item) { Add(item); }
        public void Sort(object key, bool ascending) { /* noop */ }
        public void RemoveAt(int index) { base.RemoveAt(index); }
        public new IRecord this[int index] { get => (IRecord)base[index]; set => base[index] = value; }
        public int Length => Count;
        // explicit interface
        IRecord IList<IRecord>.this[int index] { get => this[index]; set => this[index] = value; }
        public bool IList<IRecord>.IsReadOnly => false;
        int IList<IRecord>.IndexOf(IRecord item) => IndexOf(item);
        void IList<IRecord>.Insert(int index, IRecord item) { Insert(index, item); }
        void ICollection<IRecord>.Add(IRecord item) { Add(item); }
        bool ICollection<IRecord>.Contains(IRecord item) => Contains(item);
        void ICollection<IRecord>.CopyTo(IRecord[] array, int arrayIndex) { CopyTo(array, arrayIndex); }
        bool ICollection<IRecord>.Remove(IRecord item) => base.Remove(item) >= 0;
        IEnumerator<IRecord> IEnumerable<IRecord>.GetEnumerator() => GetEnumerator() as IEnumerator<IRecord>;
    }

    public static class GenericExtendedActions
    {
        public static void LogMessage(object context, string message, string moduleName)
        {
            // No-op in test context. In production, logs to OutSystems General Log.
        }
    }

    public class AppInfo
    {
        private static AppInfo _instance = new AppInfo();
        public static AppInfo GetAppInfo() => _instance;
        public object OsContext => null;
    }

    [Serializable]
    public class SStructureChanged
    {
        public SStructureChangedType EventType;
        public string AttributeName;
    }

    public enum SStructureChangedType
    {
        None = 0,
    }
}

// ---------- OutSystems.HubEdition.RuntimePlatform.Db ----------
namespace OutSystems.HubEdition.RuntimePlatform.Db
{
    public class DbAccessors
    {
        public static object GetRecord(IRecord record) => record;
    }
}

// ---------- OutSystems.HubEdition.RuntimePlatform.NewRuntime ----------
namespace OutSystems.HubEdition.RuntimePlatform.NewRuntime
{
    // marker namespace — no types needed
}

// ---------- OutSystems.Internal.Db ----------
namespace OutSystems.Internal.Db
{
    // marker namespace — no types needed
}
