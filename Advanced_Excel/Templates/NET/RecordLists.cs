using System;
using System.Data;
using System.Collections;
using System.Runtime.Serialization;
using System.Reflection;
using System.Xml;
using OutSystems.ObjectKeys;
using OutSystems.RuntimeCommon;
using OutSystems.HubEdition.RuntimePlatform;
using OutSystems.HubEdition.RuntimePlatform.Db;
using OutSystems.Internal.Db;
using OutSystems.HubEdition.RuntimePlatform.NewRuntime;

namespace OutSystems.NssAdvanced_Excel {

	/// <summary>
	/// RecordList type <code>RLCellFormatRecordList</code> that represents a record list of
	///  <code>CellFormat</code>
	/// </summary>
	[Serializable()]
	public partial class RLCellFormatRecordList: GenericRecordList<RCCellFormatRecord>, IEnumerable, IEnumerator, ISerializable {
		public static void EnsureInitialized() {}

		protected override RCCellFormatRecord GetElementDefaultValue() {
			return new RCCellFormatRecord("");
		}

		public T[] ToArray<T>(Func<RCCellFormatRecord, T> converter) {
			return ToArray(this, converter);
		}

		public static T[] ToArray<T>(RLCellFormatRecordList recordlist, Func<RCCellFormatRecord, T> converter) {
			return InnerToArray(recordlist, converter);
		}
		public static implicit operator RLCellFormatRecordList(RCCellFormatRecord[] array) {
			RLCellFormatRecordList result = new RLCellFormatRecordList();
			result.InnerFromArray(array);
			return result;
		}

		public static RLCellFormatRecordList ToList<T>(T[] array, Func <T, RCCellFormatRecord> converter) {
			RLCellFormatRecordList result = new RLCellFormatRecordList();
			result.InnerFromArray(array, converter);
			return result;
		}

		public static RLCellFormatRecordList FromRestList<T>(RestList<T> restList, Func <T, RCCellFormatRecord> converter) {
			RLCellFormatRecordList result = new RLCellFormatRecordList();
			result.InnerFromRestList(restList, converter);
			return result;
		}
		/// <summary>
		/// Default Constructor
		/// </summary>
		public RLCellFormatRecordList(): base() {
		}

		/// <summary>
		/// Constructor with transaction parameter
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLCellFormatRecordList(IDbTransaction trans): base(trans) {
		}

		/// <summary>
		/// Constructor with transaction parameter and alternate read method
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		/// <param name="alternateReadDBMethod"> Alternate Read Method</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLCellFormatRecordList(IDbTransaction trans, ReadDBMethodDelegate alternateReadDBMethod): this(trans) {
			this.alternateReadDBMethod = alternateReadDBMethod;
		}

		/// <summary>
		/// Constructor declaration for serialization
		/// </summary>
		/// <param name="info"> SerializationInfo</param>
		/// <param name="context"> StreamingContext</param>
		public RLCellFormatRecordList(SerializationInfo info, StreamingContext context): base(info, context) {
		}

		public override BitArray[] GetDefaultOptimizedValues() {
			BitArray[] def = new BitArray[1];
			def[0] = null;
			return def;
		}
		/// <summary>
		/// Create as new list
		/// </summary>
		/// <returns>The new record list</returns>
		protected override OSList<RCCellFormatRecord> NewList() {
			return new RLCellFormatRecordList();
		}


	} // RLCellFormatRecordList

	/// <summary>
	/// RecordList type <code>RLWorkbookRecordList</code> that represents a record list of
	///  <code>Workbook</code>
	/// </summary>
	[Serializable()]
	public partial class RLWorkbookRecordList: GenericRecordList<RCWorkbookRecord>, IEnumerable, IEnumerator, ISerializable {
		public static void EnsureInitialized() {}

		protected override RCWorkbookRecord GetElementDefaultValue() {
			return new RCWorkbookRecord("");
		}

		public T[] ToArray<T>(Func<RCWorkbookRecord, T> converter) {
			return ToArray(this, converter);
		}

		public static T[] ToArray<T>(RLWorkbookRecordList recordlist, Func<RCWorkbookRecord, T> converter) {
			return InnerToArray(recordlist, converter);
		}
		public static implicit operator RLWorkbookRecordList(RCWorkbookRecord[] array) {
			RLWorkbookRecordList result = new RLWorkbookRecordList();
			result.InnerFromArray(array);
			return result;
		}

		public static RLWorkbookRecordList ToList<T>(T[] array, Func <T, RCWorkbookRecord> converter) {
			RLWorkbookRecordList result = new RLWorkbookRecordList();
			result.InnerFromArray(array, converter);
			return result;
		}

		public static RLWorkbookRecordList FromRestList<T>(RestList<T> restList, Func <T, RCWorkbookRecord> converter) {
			RLWorkbookRecordList result = new RLWorkbookRecordList();
			result.InnerFromRestList(restList, converter);
			return result;
		}
		/// <summary>
		/// Default Constructor
		/// </summary>
		public RLWorkbookRecordList(): base() {
		}

		/// <summary>
		/// Constructor with transaction parameter
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLWorkbookRecordList(IDbTransaction trans): base(trans) {
		}

		/// <summary>
		/// Constructor with transaction parameter and alternate read method
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		/// <param name="alternateReadDBMethod"> Alternate Read Method</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLWorkbookRecordList(IDbTransaction trans, ReadDBMethodDelegate alternateReadDBMethod): this(trans) {
			this.alternateReadDBMethod = alternateReadDBMethod;
		}

		/// <summary>
		/// Constructor declaration for serialization
		/// </summary>
		/// <param name="info"> SerializationInfo</param>
		/// <param name="context"> StreamingContext</param>
		public RLWorkbookRecordList(SerializationInfo info, StreamingContext context): base(info, context) {
		}

		public override BitArray[] GetDefaultOptimizedValues() {
			BitArray[] def = new BitArray[1];
			def[0] = null;
			return def;
		}
		/// <summary>
		/// Create as new list
		/// </summary>
		/// <returns>The new record list</returns>
		protected override OSList<RCWorkbookRecord> NewList() {
			return new RLWorkbookRecordList();
		}


	} // RLWorkbookRecordList

	/// <summary>
	/// RecordList type <code>RLWorksheetRecordList</code> that represents a record list of
	///  <code>Worksheet</code>
	/// </summary>
	[Serializable()]
	public partial class RLWorksheetRecordList: GenericRecordList<RCWorksheetRecord>, IEnumerable, IEnumerator, ISerializable {
		public static void EnsureInitialized() {}

		protected override RCWorksheetRecord GetElementDefaultValue() {
			return new RCWorksheetRecord("");
		}

		public T[] ToArray<T>(Func<RCWorksheetRecord, T> converter) {
			return ToArray(this, converter);
		}

		public static T[] ToArray<T>(RLWorksheetRecordList recordlist, Func<RCWorksheetRecord, T> converter) {
			return InnerToArray(recordlist, converter);
		}
		public static implicit operator RLWorksheetRecordList(RCWorksheetRecord[] array) {
			RLWorksheetRecordList result = new RLWorksheetRecordList();
			result.InnerFromArray(array);
			return result;
		}

		public static RLWorksheetRecordList ToList<T>(T[] array, Func <T, RCWorksheetRecord> converter) {
			RLWorksheetRecordList result = new RLWorksheetRecordList();
			result.InnerFromArray(array, converter);
			return result;
		}

		public static RLWorksheetRecordList FromRestList<T>(RestList<T> restList, Func <T, RCWorksheetRecord> converter) {
			RLWorksheetRecordList result = new RLWorksheetRecordList();
			result.InnerFromRestList(restList, converter);
			return result;
		}
		/// <summary>
		/// Default Constructor
		/// </summary>
		public RLWorksheetRecordList(): base() {
		}

		/// <summary>
		/// Constructor with transaction parameter
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLWorksheetRecordList(IDbTransaction trans): base(trans) {
		}

		/// <summary>
		/// Constructor with transaction parameter and alternate read method
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		/// <param name="alternateReadDBMethod"> Alternate Read Method</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLWorksheetRecordList(IDbTransaction trans, ReadDBMethodDelegate alternateReadDBMethod): this(trans) {
			this.alternateReadDBMethod = alternateReadDBMethod;
		}

		/// <summary>
		/// Constructor declaration for serialization
		/// </summary>
		/// <param name="info"> SerializationInfo</param>
		/// <param name="context"> StreamingContext</param>
		public RLWorksheetRecordList(SerializationInfo info, StreamingContext context): base(info, context) {
		}

		public override BitArray[] GetDefaultOptimizedValues() {
			BitArray[] def = new BitArray[1];
			def[0] = null;
			return def;
		}
		/// <summary>
		/// Create as new list
		/// </summary>
		/// <returns>The new record list</returns>
		protected override OSList<RCWorksheetRecord> NewList() {
			return new RLWorksheetRecordList();
		}


	} // RLWorksheetRecordList

	/// <summary>
	/// RecordList type <code>RLColorRecordList</code> that represents a record list of <code>Color</code>
	/// </summary>
	[Serializable()]
	public partial class RLColorRecordList: GenericRecordList<RCColorRecord>, IEnumerable, IEnumerator, ISerializable {
		public static void EnsureInitialized() {}

		protected override RCColorRecord GetElementDefaultValue() {
			return new RCColorRecord("");
		}

		public T[] ToArray<T>(Func<RCColorRecord, T> converter) {
			return ToArray(this, converter);
		}

		public static T[] ToArray<T>(RLColorRecordList recordlist, Func<RCColorRecord, T> converter) {
			return InnerToArray(recordlist, converter);
		}
		public static implicit operator RLColorRecordList(RCColorRecord[] array) {
			RLColorRecordList result = new RLColorRecordList();
			result.InnerFromArray(array);
			return result;
		}

		public static RLColorRecordList ToList<T>(T[] array, Func <T, RCColorRecord> converter) {
			RLColorRecordList result = new RLColorRecordList();
			result.InnerFromArray(array, converter);
			return result;
		}

		public static RLColorRecordList FromRestList<T>(RestList<T> restList, Func <T, RCColorRecord> converter) {
			RLColorRecordList result = new RLColorRecordList();
			result.InnerFromRestList(restList, converter);
			return result;
		}
		/// <summary>
		/// Default Constructor
		/// </summary>
		public RLColorRecordList(): base() {
		}

		/// <summary>
		/// Constructor with transaction parameter
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLColorRecordList(IDbTransaction trans): base(trans) {
		}

		/// <summary>
		/// Constructor with transaction parameter and alternate read method
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		/// <param name="alternateReadDBMethod"> Alternate Read Method</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLColorRecordList(IDbTransaction trans, ReadDBMethodDelegate alternateReadDBMethod): this(trans) {
			this.alternateReadDBMethod = alternateReadDBMethod;
		}

		/// <summary>
		/// Constructor declaration for serialization
		/// </summary>
		/// <param name="info"> SerializationInfo</param>
		/// <param name="context"> StreamingContext</param>
		public RLColorRecordList(SerializationInfo info, StreamingContext context): base(info, context) {
		}

		public override BitArray[] GetDefaultOptimizedValues() {
			BitArray[] def = new BitArray[1];
			def[0] = null;
			return def;
		}
		/// <summary>
		/// Create as new list
		/// </summary>
		/// <returns>The new record list</returns>
		protected override OSList<RCColorRecord> NewList() {
			return new RLColorRecordList();
		}


	} // RLColorRecordList

	/// <summary>
	/// RecordList type <code>RLDimensionRecordList</code> that represents a record list of
	///  <code>Dimension</code>
	/// </summary>
	[Serializable()]
	public partial class RLDimensionRecordList: GenericRecordList<RCDimensionRecord>, IEnumerable, IEnumerator, ISerializable {
		public static void EnsureInitialized() {}

		protected override RCDimensionRecord GetElementDefaultValue() {
			return new RCDimensionRecord("");
		}

		public T[] ToArray<T>(Func<RCDimensionRecord, T> converter) {
			return ToArray(this, converter);
		}

		public static T[] ToArray<T>(RLDimensionRecordList recordlist, Func<RCDimensionRecord, T> converter) {
			return InnerToArray(recordlist, converter);
		}
		public static implicit operator RLDimensionRecordList(RCDimensionRecord[] array) {
			RLDimensionRecordList result = new RLDimensionRecordList();
			result.InnerFromArray(array);
			return result;
		}

		public static RLDimensionRecordList ToList<T>(T[] array, Func <T, RCDimensionRecord> converter) {
			RLDimensionRecordList result = new RLDimensionRecordList();
			result.InnerFromArray(array, converter);
			return result;
		}

		public static RLDimensionRecordList FromRestList<T>(RestList<T> restList, Func <T, RCDimensionRecord> converter) {
			RLDimensionRecordList result = new RLDimensionRecordList();
			result.InnerFromRestList(restList, converter);
			return result;
		}
		/// <summary>
		/// Default Constructor
		/// </summary>
		public RLDimensionRecordList(): base() {
		}

		/// <summary>
		/// Constructor with transaction parameter
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLDimensionRecordList(IDbTransaction trans): base(trans) {
		}

		/// <summary>
		/// Constructor with transaction parameter and alternate read method
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		/// <param name="alternateReadDBMethod"> Alternate Read Method</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLDimensionRecordList(IDbTransaction trans, ReadDBMethodDelegate alternateReadDBMethod): this(trans) {
			this.alternateReadDBMethod = alternateReadDBMethod;
		}

		/// <summary>
		/// Constructor declaration for serialization
		/// </summary>
		/// <param name="info"> SerializationInfo</param>
		/// <param name="context"> StreamingContext</param>
		public RLDimensionRecordList(SerializationInfo info, StreamingContext context): base(info, context) {
		}

		public override BitArray[] GetDefaultOptimizedValues() {
			BitArray[] def = new BitArray[1];
			def[0] = null;
			return def;
		}
		/// <summary>
		/// Create as new list
		/// </summary>
		/// <returns>The new record list</returns>
		protected override OSList<RCDimensionRecord> NewList() {
			return new RLDimensionRecordList();
		}


	} // RLDimensionRecordList

	/// <summary>
	/// RecordList type <code>RLAddressRecordList</code> that represents a record list of
	///  <code>Address</code>
	/// </summary>
	[Serializable()]
	public partial class RLAddressRecordList: GenericRecordList<RCAddressRecord>, IEnumerable, IEnumerator, ISerializable {
		public static void EnsureInitialized() {}

		protected override RCAddressRecord GetElementDefaultValue() {
			return new RCAddressRecord("");
		}

		public T[] ToArray<T>(Func<RCAddressRecord, T> converter) {
			return ToArray(this, converter);
		}

		public static T[] ToArray<T>(RLAddressRecordList recordlist, Func<RCAddressRecord, T> converter) {
			return InnerToArray(recordlist, converter);
		}
		public static implicit operator RLAddressRecordList(RCAddressRecord[] array) {
			RLAddressRecordList result = new RLAddressRecordList();
			result.InnerFromArray(array);
			return result;
		}

		public static RLAddressRecordList ToList<T>(T[] array, Func <T, RCAddressRecord> converter) {
			RLAddressRecordList result = new RLAddressRecordList();
			result.InnerFromArray(array, converter);
			return result;
		}

		public static RLAddressRecordList FromRestList<T>(RestList<T> restList, Func <T, RCAddressRecord> converter) {
			RLAddressRecordList result = new RLAddressRecordList();
			result.InnerFromRestList(restList, converter);
			return result;
		}
		/// <summary>
		/// Default Constructor
		/// </summary>
		public RLAddressRecordList(): base() {
		}

		/// <summary>
		/// Constructor with transaction parameter
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLAddressRecordList(IDbTransaction trans): base(trans) {
		}

		/// <summary>
		/// Constructor with transaction parameter and alternate read method
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		/// <param name="alternateReadDBMethod"> Alternate Read Method</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLAddressRecordList(IDbTransaction trans, ReadDBMethodDelegate alternateReadDBMethod): this(trans) {
			this.alternateReadDBMethod = alternateReadDBMethod;
		}

		/// <summary>
		/// Constructor declaration for serialization
		/// </summary>
		/// <param name="info"> SerializationInfo</param>
		/// <param name="context"> StreamingContext</param>
		public RLAddressRecordList(SerializationInfo info, StreamingContext context): base(info, context) {
		}

		public override BitArray[] GetDefaultOptimizedValues() {
			BitArray[] def = new BitArray[1];
			def[0] = null;
			return def;
		}
		/// <summary>
		/// Create as new list
		/// </summary>
		/// <returns>The new record list</returns>
		protected override OSList<RCAddressRecord> NewList() {
			return new RLAddressRecordList();
		}


	} // RLAddressRecordList

	/// <summary>
	/// RecordList type <code>RLRangeRecordList</code> that represents a record list of <code>Range</code>
	/// </summary>
	[Serializable()]
	public partial class RLRangeRecordList: GenericRecordList<RCRangeRecord>, IEnumerable, IEnumerator, ISerializable {
		public static void EnsureInitialized() {}

		protected override RCRangeRecord GetElementDefaultValue() {
			return new RCRangeRecord("");
		}

		public T[] ToArray<T>(Func<RCRangeRecord, T> converter) {
			return ToArray(this, converter);
		}

		public static T[] ToArray<T>(RLRangeRecordList recordlist, Func<RCRangeRecord, T> converter) {
			return InnerToArray(recordlist, converter);
		}
		public static implicit operator RLRangeRecordList(RCRangeRecord[] array) {
			RLRangeRecordList result = new RLRangeRecordList();
			result.InnerFromArray(array);
			return result;
		}

		public static RLRangeRecordList ToList<T>(T[] array, Func <T, RCRangeRecord> converter) {
			RLRangeRecordList result = new RLRangeRecordList();
			result.InnerFromArray(array, converter);
			return result;
		}

		public static RLRangeRecordList FromRestList<T>(RestList<T> restList, Func <T, RCRangeRecord> converter) {
			RLRangeRecordList result = new RLRangeRecordList();
			result.InnerFromRestList(restList, converter);
			return result;
		}
		/// <summary>
		/// Default Constructor
		/// </summary>
		public RLRangeRecordList(): base() {
		}

		/// <summary>
		/// Constructor with transaction parameter
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLRangeRecordList(IDbTransaction trans): base(trans) {
		}

		/// <summary>
		/// Constructor with transaction parameter and alternate read method
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		/// <param name="alternateReadDBMethod"> Alternate Read Method</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLRangeRecordList(IDbTransaction trans, ReadDBMethodDelegate alternateReadDBMethod): this(trans) {
			this.alternateReadDBMethod = alternateReadDBMethod;
		}

		/// <summary>
		/// Constructor declaration for serialization
		/// </summary>
		/// <param name="info"> SerializationInfo</param>
		/// <param name="context"> StreamingContext</param>
		public RLRangeRecordList(SerializationInfo info, StreamingContext context): base(info, context) {
		}

		public override BitArray[] GetDefaultOptimizedValues() {
			BitArray[] def = new BitArray[1];
			def[0] = null;
			return def;
		}
		/// <summary>
		/// Create as new list
		/// </summary>
		/// <returns>The new record list</returns>
		protected override OSList<RCRangeRecord> NewList() {
			return new RLRangeRecordList();
		}


	} // RLRangeRecordList

	/// <summary>
	/// RecordList type <code>RLDataSeriesRecordList</code> that represents a record list of
	///  <code>DataSeries</code>
	/// </summary>
	[Serializable()]
	public partial class RLDataSeriesRecordList: GenericRecordList<RCDataSeriesRecord>, IEnumerable, IEnumerator, ISerializable {
		public static void EnsureInitialized() {}

		protected override RCDataSeriesRecord GetElementDefaultValue() {
			return new RCDataSeriesRecord("");
		}

		public T[] ToArray<T>(Func<RCDataSeriesRecord, T> converter) {
			return ToArray(this, converter);
		}

		public static T[] ToArray<T>(RLDataSeriesRecordList recordlist, Func<RCDataSeriesRecord, T> converter) {
			return InnerToArray(recordlist, converter);
		}
		public static implicit operator RLDataSeriesRecordList(RCDataSeriesRecord[] array) {
			RLDataSeriesRecordList result = new RLDataSeriesRecordList();
			result.InnerFromArray(array);
			return result;
		}

		public static RLDataSeriesRecordList ToList<T>(T[] array, Func <T, RCDataSeriesRecord> converter) {
			RLDataSeriesRecordList result = new RLDataSeriesRecordList();
			result.InnerFromArray(array, converter);
			return result;
		}

		public static RLDataSeriesRecordList FromRestList<T>(RestList<T> restList, Func <T, RCDataSeriesRecord> converter) {
			RLDataSeriesRecordList result = new RLDataSeriesRecordList();
			result.InnerFromRestList(restList, converter);
			return result;
		}
		/// <summary>
		/// Default Constructor
		/// </summary>
		public RLDataSeriesRecordList(): base() {
		}

		/// <summary>
		/// Constructor with transaction parameter
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLDataSeriesRecordList(IDbTransaction trans): base(trans) {
		}

		/// <summary>
		/// Constructor with transaction parameter and alternate read method
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		/// <param name="alternateReadDBMethod"> Alternate Read Method</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLDataSeriesRecordList(IDbTransaction trans, ReadDBMethodDelegate alternateReadDBMethod): this(trans) {
			this.alternateReadDBMethod = alternateReadDBMethod;
		}

		/// <summary>
		/// Constructor declaration for serialization
		/// </summary>
		/// <param name="info"> SerializationInfo</param>
		/// <param name="context"> StreamingContext</param>
		public RLDataSeriesRecordList(SerializationInfo info, StreamingContext context): base(info, context) {
		}

		public override BitArray[] GetDefaultOptimizedValues() {
			BitArray[] def = new BitArray[1];
			def[0] = null;
			return def;
		}
		/// <summary>
		/// Create as new list
		/// </summary>
		/// <returns>The new record list</returns>
		protected override OSList<RCDataSeriesRecord> NewList() {
			return new RLDataSeriesRecordList();
		}


	} // RLDataSeriesRecordList
}
