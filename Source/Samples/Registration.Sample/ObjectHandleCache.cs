using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using ExcelDna.Integration;
using ExcelDna.Registration;
using System.Linq.Expressions;

namespace Registration.Sample
{
    /// <summary>
    /// Any types decorated with this attribute will be marshalled to Excel by reference through the ParameterConversionConfiguration.ReferenceMashaller object
    /// </summary>
    [AttributeUsage(AttributeTargets.Class|AttributeTargets.Interface, Inherited = true)]
    public class ExcelMarshalByHandleAttribute : Attribute
    {
        public ExcelMarshalByHandleAttribute()
        {
        }
    }

    /// <summary>
    /// This class implements an object cache to marshall object identities back and forth the Excel workbook.
    /// It was designed as a singleton because it uses thread local storage to support multi-threaded scenarios in an efficient way.
    /// </summary>
    internal class ObjectHandleCache : FunctionExecutionHandler
    {
        private static volatile ObjectHandleCache _instance;
        private static readonly object SyncRoot = new Object();

        private ObjectHandleCache() { }

        /// <summary>
        /// The singleton instance of this class
        /// </summary>
        public static ObjectHandleCache Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (SyncRoot)
                    {
                        if (_instance == null)
                            _instance = new ObjectHandleCache();
                    }
                }

                return _instance;
            }
        }

        private ReaderWriterLockSlim _rwLock = new ReaderWriterLockSlim();

        /// <summary>
        /// This tracks which object id's have been create din which cell
        /// </summary>
        private Dictionary<ExcelReference, HashSet<int>> BirthCell = new Dictionary<ExcelReference, HashSet<int>>();

        /// <summary>
        /// This is where we actually look objects up
        /// </summary>
        private Dictionary<int, object> IdLookup = new Dictionary<int, object>();

        /// <summary>
        /// This acculates all objects created during the possibly nested evaluations that happen in a cell.
        /// It is cleared every time the thread sees a different cell.
        /// </summary>
        [ThreadStatic]
        private static Dictionary<int, object> TempObjects = new Dictionary<int, object>();

        /// <summary>
        /// This is the cell being handled by the current thread
        /// </summary>
        [ThreadStatic]
        private static ExcelReference _currentCell;

        /// <summary>
        /// An atomically incremented counted that provides the object identifiers.
        /// </summary>
        private static int _idCounter = 0;

        /// NullXlRef is what we get for calls from VBA
        private static readonly ExcelReference NullXlRef = new ExcelReference(-1, -1);

        private static readonly char Separator = '@';

        /// <summary>
        /// This method is called to look up an object id string in the cache.
        /// </summary>
        /// <param name="idString">The object id string coming from Excel</param>
        /// <returns>The object or null</returns>
        public object Lookup(string idString)
        {
            int idPos = idString.IndexOf('@');
            if (idPos<0)
                throw new ArgumentException($"Object id '{idString}' is not in the format <name>{Separator}<id>");
            int id = int.Parse(idString.Substring(idPos + 1));
            object result;
            if (!TempObjects.TryGetValue(id, out result))
            {
                _rwLock.EnterReadLock();
                try
                {
                    IdLookup.TryGetValue(id, out result);
                }
                finally
                {
                    _rwLock.ExitReadLock();
                }
            }
            return result;
        }

        public T Lookup<T>(string idString) where T : class
        {
            object o = Lookup(idString);
            T r = o as T;
            if (r == null)
                throw new ArgumentException($"Object '{idString}' is not of type {typeof(T)}, it is of type {o.GetType()}.");
            return r;
        }

        /// <summary>
        /// This method is called to store any objects returned by a function.
        /// </summary>
        /// <param name="o">The object</param>
        /// <returns>The string id</returns>
        public string Store(object o)
        {
            int id = Interlocked.Increment(ref _idCounter);
            TempObjects[id] = o;
            _rwLock.EnterWriteLock();
            try
            {
                // The following line assumes that SetCurrentCell has been called before getting here
                HashSet<int> objectsBornInTheCurrentCell = BirthCell[_currentCell];
                objectsBornInTheCurrentCell.Add(id);
                IdLookup[id] = o;
            }
            finally
            {
                _rwLock.ExitWriteLock();
            }
            return $"{o.GetType().Name}{Separator}{id}";
        }

        public string Store<T>(T o) where T : class
        {
            return Store((object)o);
        }

        /// <summary>
        /// This method is called to tell the cache that we are starting to evaluate a new cell.
        /// </summary>
        public void SetCurrentCell()
        {
            ExcelReference previousCell = _currentCell;
            ExcelReference thisCell = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            thisCell = thisCell ?? NullXlRef;
            _rwLock.EnterWriteLock();
            try
            {
                // If we are now processing a different cell, then we get rid of any old temp objects
                if (thisCell != previousCell)
                {
                    HashSet<int> objectsCreatedAtThePreviousCell = null;
                    if (previousCell != null)
                        BirthCell.TryGetValue(previousCell, out objectsCreatedAtThePreviousCell);
                    foreach (int tempId in TempObjects.Keys)
                        if (objectsCreatedAtThePreviousCell == null || !objectsCreatedAtThePreviousCell.Contains(tempId))
                            IdLookup.Remove(tempId);
                    TempObjects.Clear();
                    _currentCell = thisCell;
                }

                HashSet<int> objectsCreatedAtTheCurrentCell = null;
                BirthCell.TryGetValue(thisCell, out objectsCreatedAtTheCurrentCell);

                if (objectsCreatedAtTheCurrentCell == null)
                {
                    objectsCreatedAtTheCurrentCell = new HashSet<int>();
                    BirthCell[thisCell] = objectsCreatedAtTheCurrentCell;
                }
                else
                {
                    foreach (var id in objectsCreatedAtTheCurrentCell)
                        IdLookup.Remove(id);
                    objectsCreatedAtTheCurrentCell.Clear();
                }
            }
            finally
            {
                _rwLock.ExitWriteLock();
            }
        }

        [ThreadStatic]
        private static bool _threadIsReturningACollection;

        public override void OnEntry(FunctionExecutionArgs args)
        {
            _threadIsReturningACollection = true;
            ObjectHandleCache.Instance.SetCurrentCell();
        }

        public override void OnExit(FunctionExecutionArgs args)
        {
            _threadIsReturningACollection = false;
        }

        internal static FunctionExecutionHandler ObjectHandleTrackingSelector(ExcelFunctionRegistration functionRegistration)
        {
            FunctionExecutionHandler result = null;
            if (functionRegistration.FunctionAttribute.GetType() != typeof(ExcelMapArrayFunctionAttribute))
                result = Instance;
            return result;
        }
    }

    internal static class ParameterConversionConfigurationEntensions
    {
        static ParameterConversionConfiguration AddHandleConversion(this ParameterConversionConfiguration config,
            Type targetType)
        {
            if (targetType.IsValueType)
                throw new ArgumentException("Only reference types can be marshalled to/from XL as object handles.");

            #region register parameter lambda

            {
                var idString = Expression.Parameter(typeof(string), "idString");
                var paramLambda = Expression.Lambda(
                    Expression.Call(Expression.Constant(ObjectHandleCache.Instance), "Lookup", new Type[] {targetType},
                        idString),
                    idString);
                config.AddParameterConversion((Type x, ExcelParameterRegistration reg) => paramLambda, targetType);
            }

            #endregion

            #region Register return lambda

            {
                var inputObject = Expression.Parameter(targetType, "inputObject");
                var returnLambda = Expression.Lambda(
                    Expression.Call(Expression.Constant(ObjectHandleCache.Instance), "Store",
                        new Type[] {targetType},
                        inputObject),
                    inputObject);
                config.AddReturnConversion((Type t, ExcelReturnRegistration reg) => returnLambda, targetType, true);
            }

            #endregion

            return config;
        }

        public static ParameterConversionConfiguration AddHandleConversions(this ParameterConversionConfiguration config)
        {
            var types = from typ in Assembly.GetCallingAssembly().GetTypes()
                where typ.GetCustomAttribute<ExcelMarshalByHandleAttribute>() != null
                select typ;
            foreach (Type t in types)
                config.AddHandleConversion(t);
            return config;
        }
    }
}
