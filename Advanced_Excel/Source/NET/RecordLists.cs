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

	/// <summary>
	/// RecordList type <code>RLConditionalFormatItemRecordList</code> that represents a record list of
	///  <code>ConditionalFormatItem</code>
	/// </summary>
	[Serializable()]
	public partial class RLConditionalFormatItemRecordList: GenericRecordList<RCConditionalFormatItemRecord>, IEnumerable, IEnumerator, ISerializable {
		public static void EnsureInitialized() {}

		protected override RCConditionalFormatItemRecord GetElementDefaultValue() {
			return new RCConditionalFormatItemRecord("");
		}

		public T[] ToArray<T>(Func<RCConditionalFormatItemRecord, T> converter) {
			return ToArray(this, converter);
		}

		public static T[] ToArray<T>(RLConditionalFormatItemRecordList recordlist, Func<RCConditionalFormatItemRecord, T> converter) {
			return InnerToArray(recordlist, converter);
		}
		public static implicit operator RLConditionalFormatItemRecordList(RCConditionalFormatItemRecord[] array) {
			RLConditionalFormatItemRecordList result = new RLConditionalFormatItemRecordList();
			result.InnerFromArray(array);
			return result;
		}

		public static RLConditionalFormatItemRecordList ToList<T>(T[] array, Func <T, RCConditionalFormatItemRecord> converter) {
			RLConditionalFormatItemRecordList result = new RLConditionalFormatItemRecordList();
			result.InnerFromArray(array, converter);
			return result;
		}

		public static RLConditionalFormatItemRecordList FromRestList<T>(RestList<T> restList, Func <T, RCConditionalFormatItemRecord> converter) {
			RLConditionalFormatItemRecordList result = new RLConditionalFormatItemRecordList();
			result.InnerFromRestList(restList, converter);
			return result;
		}
		/// <summary>
		/// Default Constructor
		/// </summary>
		public RLConditionalFormatItemRecordList(): base() {
		}

		/// <summary>
		/// Constructor with transaction parameter
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLConditionalFormatItemRecordList(IDbTransaction trans): base(trans) {
		}

		/// <summary>
		/// Constructor with transaction parameter and alternate read method
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		/// <param name="alternateReadDBMethod"> Alternate Read Method</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLConditionalFormatItemRecordList(IDbTransaction trans, ReadDBMethodDelegate alternateReadDBMethod): this(trans) {
			this.alternateReadDBMethod = alternateReadDBMethod;
		}

		/// <summary>
		/// Constructor declaration for serialization
		/// </summary>
		/// <param name="info"> SerializationInfo</param>
		/// <param name="context"> StreamingContext</param>
		public RLConditionalFormatItemRecordList(SerializationInfo info, StreamingContext context): base(info, context) {
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
		protected override OSList<RCConditionalFormatItemRecord> NewList() {
			return new RLConditionalFormatItemRecordList();
		}


	} // RLConditionalFormatItemRecordList

	/// <summary>
	/// RecordList type <code>RLConditionalFormatStyleRecordList</code> that represents a record list of
	///  <code>ConditionalFormatStyle</code>
	/// </summary>
	[Serializable()]
	public partial class RLConditionalFormatStyleRecordList: GenericRecordList<RCConditionalFormatStyleRecord>, IEnumerable, IEnumerator, ISerializable {
		public static void EnsureInitialized() {}

		protected override RCConditionalFormatStyleRecord GetElementDefaultValue() {
			return new RCConditionalFormatStyleRecord("");
		}

		public T[] ToArray<T>(Func<RCConditionalFormatStyleRecord, T> converter) {
			return ToArray(this, converter);
		}

		public static T[] ToArray<T>(RLConditionalFormatStyleRecordList recordlist, Func<RCConditionalFormatStyleRecord, T> converter) {
			return InnerToArray(recordlist, converter);
		}
		public static implicit operator RLConditionalFormatStyleRecordList(RCConditionalFormatStyleRecord[] array) {
			RLConditionalFormatStyleRecordList result = new RLConditionalFormatStyleRecordList();
			result.InnerFromArray(array);
			return result;
		}

		public static RLConditionalFormatStyleRecordList ToList<T>(T[] array, Func <T, RCConditionalFormatStyleRecord> converter) {
			RLConditionalFormatStyleRecordList result = new RLConditionalFormatStyleRecordList();
			result.InnerFromArray(array, converter);
			return result;
		}

		public static RLConditionalFormatStyleRecordList FromRestList<T>(RestList<T> restList, Func <T, RCConditionalFormatStyleRecord> converter) {
			RLConditionalFormatStyleRecordList result = new RLConditionalFormatStyleRecordList();
			result.InnerFromRestList(restList, converter);
			return result;
		}
		/// <summary>
		/// Default Constructor
		/// </summary>
		public RLConditionalFormatStyleRecordList(): base() {
		}

		/// <summary>
		/// Constructor with transaction parameter
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLConditionalFormatStyleRecordList(IDbTransaction trans): base(trans) {
		}

		/// <summary>
		/// Constructor with transaction parameter and alternate read method
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		/// <param name="alternateReadDBMethod"> Alternate Read Method</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLConditionalFormatStyleRecordList(IDbTransaction trans, ReadDBMethodDelegate alternateReadDBMethod): this(trans) {
			this.alternateReadDBMethod = alternateReadDBMethod;
		}

		/// <summary>
		/// Constructor declaration for serialization
		/// </summary>
		/// <param name="info"> SerializationInfo</param>
		/// <param name="context"> StreamingContext</param>
		public RLConditionalFormatStyleRecordList(SerializationInfo info, StreamingContext context): base(info, context) {
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
		protected override OSList<RCConditionalFormatStyleRecord> NewList() {
			return new RLConditionalFormatStyleRecordList();
		}


	} // RLConditionalFormatStyleRecordList

	/// <summary>
	/// RecordList type <code>RLBorderStyleRecordList</code> that represents a record list of
	///  <code>BorderStyle</code>
	/// </summary>
	[Serializable()]
	public partial class RLBorderStyleRecordList: GenericRecordList<RCBorderStyleRecord>, IEnumerable, IEnumerator, ISerializable {
		public static void EnsureInitialized() {}

		protected override RCBorderStyleRecord GetElementDefaultValue() {
			return new RCBorderStyleRecord("");
		}

		public T[] ToArray<T>(Func<RCBorderStyleRecord, T> converter) {
			return ToArray(this, converter);
		}

		public static T[] ToArray<T>(RLBorderStyleRecordList recordlist, Func<RCBorderStyleRecord, T> converter) {
			return InnerToArray(recordlist, converter);
		}
		public static implicit operator RLBorderStyleRecordList(RCBorderStyleRecord[] array) {
			RLBorderStyleRecordList result = new RLBorderStyleRecordList();
			result.InnerFromArray(array);
			return result;
		}

		public static RLBorderStyleRecordList ToList<T>(T[] array, Func <T, RCBorderStyleRecord> converter) {
			RLBorderStyleRecordList result = new RLBorderStyleRecordList();
			result.InnerFromArray(array, converter);
			return result;
		}

		public static RLBorderStyleRecordList FromRestList<T>(RestList<T> restList, Func <T, RCBorderStyleRecord> converter) {
			RLBorderStyleRecordList result = new RLBorderStyleRecordList();
			result.InnerFromRestList(restList, converter);
			return result;
		}
		/// <summary>
		/// Default Constructor
		/// </summary>
		public RLBorderStyleRecordList(): base() {
		}

		/// <summary>
		/// Constructor with transaction parameter
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLBorderStyleRecordList(IDbTransaction trans): base(trans) {
		}

		/// <summary>
		/// Constructor with transaction parameter and alternate read method
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		/// <param name="alternateReadDBMethod"> Alternate Read Method</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLBorderStyleRecordList(IDbTransaction trans, ReadDBMethodDelegate alternateReadDBMethod): this(trans) {
			this.alternateReadDBMethod = alternateReadDBMethod;
		}

		/// <summary>
		/// Constructor declaration for serialization
		/// </summary>
		/// <param name="info"> SerializationInfo</param>
		/// <param name="context"> StreamingContext</param>
		public RLBorderStyleRecordList(SerializationInfo info, StreamingContext context): base(info, context) {
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
		protected override OSList<RCBorderStyleRecord> NewList() {
			return new RLBorderStyleRecordList();
		}


	} // RLBorderStyleRecordList

	/// <summary>
	/// RecordList type <code>RLFillStyleRecordList</code> that represents a record list of
	///  <code>FillStyle</code>
	/// </summary>
	[Serializable()]
	public partial class RLFillStyleRecordList: GenericRecordList<RCFillStyleRecord>, IEnumerable, IEnumerator, ISerializable {
		public static void EnsureInitialized() {}

		protected override RCFillStyleRecord GetElementDefaultValue() {
			return new RCFillStyleRecord("");
		}

		public T[] ToArray<T>(Func<RCFillStyleRecord, T> converter) {
			return ToArray(this, converter);
		}

		public static T[] ToArray<T>(RLFillStyleRecordList recordlist, Func<RCFillStyleRecord, T> converter) {
			return InnerToArray(recordlist, converter);
		}
		public static implicit operator RLFillStyleRecordList(RCFillStyleRecord[] array) {
			RLFillStyleRecordList result = new RLFillStyleRecordList();
			result.InnerFromArray(array);
			return result;
		}

		public static RLFillStyleRecordList ToList<T>(T[] array, Func <T, RCFillStyleRecord> converter) {
			RLFillStyleRecordList result = new RLFillStyleRecordList();
			result.InnerFromArray(array, converter);
			return result;
		}

		public static RLFillStyleRecordList FromRestList<T>(RestList<T> restList, Func <T, RCFillStyleRecord> converter) {
			RLFillStyleRecordList result = new RLFillStyleRecordList();
			result.InnerFromRestList(restList, converter);
			return result;
		}
		/// <summary>
		/// Default Constructor
		/// </summary>
		public RLFillStyleRecordList(): base() {
		}

		/// <summary>
		/// Constructor with transaction parameter
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLFillStyleRecordList(IDbTransaction trans): base(trans) {
		}

		/// <summary>
		/// Constructor with transaction parameter and alternate read method
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		/// <param name="alternateReadDBMethod"> Alternate Read Method</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLFillStyleRecordList(IDbTransaction trans, ReadDBMethodDelegate alternateReadDBMethod): this(trans) {
			this.alternateReadDBMethod = alternateReadDBMethod;
		}

		/// <summary>
		/// Constructor declaration for serialization
		/// </summary>
		/// <param name="info"> SerializationInfo</param>
		/// <param name="context"> StreamingContext</param>
		public RLFillStyleRecordList(SerializationInfo info, StreamingContext context): base(info, context) {
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
		protected override OSList<RCFillStyleRecord> NewList() {
			return new RLFillStyleRecordList();
		}


	} // RLFillStyleRecordList

	/// <summary>
	/// RecordList type <code>RLFontStyleRecordList</code> that represents a record list of
	///  <code>FontStyle</code>
	/// </summary>
	[Serializable()]
	public partial class RLFontStyleRecordList: GenericRecordList<RCFontStyleRecord>, IEnumerable, IEnumerator, ISerializable {
		public static void EnsureInitialized() {}

		protected override RCFontStyleRecord GetElementDefaultValue() {
			return new RCFontStyleRecord("");
		}

		public T[] ToArray<T>(Func<RCFontStyleRecord, T> converter) {
			return ToArray(this, converter);
		}

		public static T[] ToArray<T>(RLFontStyleRecordList recordlist, Func<RCFontStyleRecord, T> converter) {
			return InnerToArray(recordlist, converter);
		}
		public static implicit operator RLFontStyleRecordList(RCFontStyleRecord[] array) {
			RLFontStyleRecordList result = new RLFontStyleRecordList();
			result.InnerFromArray(array);
			return result;
		}

		public static RLFontStyleRecordList ToList<T>(T[] array, Func <T, RCFontStyleRecord> converter) {
			RLFontStyleRecordList result = new RLFontStyleRecordList();
			result.InnerFromArray(array, converter);
			return result;
		}

		public static RLFontStyleRecordList FromRestList<T>(RestList<T> restList, Func <T, RCFontStyleRecord> converter) {
			RLFontStyleRecordList result = new RLFontStyleRecordList();
			result.InnerFromRestList(restList, converter);
			return result;
		}
		/// <summary>
		/// Default Constructor
		/// </summary>
		public RLFontStyleRecordList(): base() {
		}

		/// <summary>
		/// Constructor with transaction parameter
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLFontStyleRecordList(IDbTransaction trans): base(trans) {
		}

		/// <summary>
		/// Constructor with transaction parameter and alternate read method
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		/// <param name="alternateReadDBMethod"> Alternate Read Method</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLFontStyleRecordList(IDbTransaction trans, ReadDBMethodDelegate alternateReadDBMethod): this(trans) {
			this.alternateReadDBMethod = alternateReadDBMethod;
		}

		/// <summary>
		/// Constructor declaration for serialization
		/// </summary>
		/// <param name="info"> SerializationInfo</param>
		/// <param name="context"> StreamingContext</param>
		public RLFontStyleRecordList(SerializationInfo info, StreamingContext context): base(info, context) {
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
		protected override OSList<RCFontStyleRecord> NewList() {
			return new RLFontStyleRecordList();
		}


	} // RLFontStyleRecordList

	/// <summary>
	/// RecordList type <code>RLNewSheetRecordList</code> that represents a record list of
	///  <code>NewSheet</code>
	/// </summary>
	[Serializable()]
	public partial class RLNewSheetRecordList: GenericRecordList<RCNewSheetRecord>, IEnumerable, IEnumerator, ISerializable {
		public static void EnsureInitialized() {}

		protected override RCNewSheetRecord GetElementDefaultValue() {
			return new RCNewSheetRecord("");
		}

		public T[] ToArray<T>(Func<RCNewSheetRecord, T> converter) {
			return ToArray(this, converter);
		}

		public static T[] ToArray<T>(RLNewSheetRecordList recordlist, Func<RCNewSheetRecord, T> converter) {
			return InnerToArray(recordlist, converter);
		}
		public static implicit operator RLNewSheetRecordList(RCNewSheetRecord[] array) {
			RLNewSheetRecordList result = new RLNewSheetRecordList();
			result.InnerFromArray(array);
			return result;
		}

		public static RLNewSheetRecordList ToList<T>(T[] array, Func <T, RCNewSheetRecord> converter) {
			RLNewSheetRecordList result = new RLNewSheetRecordList();
			result.InnerFromArray(array, converter);
			return result;
		}

		public static RLNewSheetRecordList FromRestList<T>(RestList<T> restList, Func <T, RCNewSheetRecord> converter) {
			RLNewSheetRecordList result = new RLNewSheetRecordList();
			result.InnerFromRestList(restList, converter);
			return result;
		}
		/// <summary>
		/// Default Constructor
		/// </summary>
		public RLNewSheetRecordList(): base() {
		}

		/// <summary>
		/// Constructor with transaction parameter
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLNewSheetRecordList(IDbTransaction trans): base(trans) {
		}

		/// <summary>
		/// Constructor with transaction parameter and alternate read method
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		/// <param name="alternateReadDBMethod"> Alternate Read Method</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLNewSheetRecordList(IDbTransaction trans, ReadDBMethodDelegate alternateReadDBMethod): this(trans) {
			this.alternateReadDBMethod = alternateReadDBMethod;
		}

		/// <summary>
		/// Constructor declaration for serialization
		/// </summary>
		/// <param name="info"> SerializationInfo</param>
		/// <param name="context"> StreamingContext</param>
		public RLNewSheetRecordList(SerializationInfo info, StreamingContext context): base(info, context) {
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
		protected override OSList<RCNewSheetRecord> NewList() {
			return new RLNewSheetRecordList();
		}


	} // RLNewSheetRecordList

	/// <summary>
	/// RecordList type <code>RLImageRecordList</code> that represents a record list of <code>Image</code>
	/// </summary>
	[Serializable()]
	public partial class RLImageRecordList: GenericRecordList<RCImageRecord>, IEnumerable, IEnumerator, ISerializable {
		public static void EnsureInitialized() {}

		protected override RCImageRecord GetElementDefaultValue() {
			return new RCImageRecord("");
		}

		public T[] ToArray<T>(Func<RCImageRecord, T> converter) {
			return ToArray(this, converter);
		}

		public static T[] ToArray<T>(RLImageRecordList recordlist, Func<RCImageRecord, T> converter) {
			return InnerToArray(recordlist, converter);
		}
		public static implicit operator RLImageRecordList(RCImageRecord[] array) {
			RLImageRecordList result = new RLImageRecordList();
			result.InnerFromArray(array);
			return result;
		}

		public static RLImageRecordList ToList<T>(T[] array, Func <T, RCImageRecord> converter) {
			RLImageRecordList result = new RLImageRecordList();
			result.InnerFromArray(array, converter);
			return result;
		}

		public static RLImageRecordList FromRestList<T>(RestList<T> restList, Func <T, RCImageRecord> converter) {
			RLImageRecordList result = new RLImageRecordList();
			result.InnerFromRestList(restList, converter);
			return result;
		}
		/// <summary>
		/// Default Constructor
		/// </summary>
		public RLImageRecordList(): base() {
		}

		/// <summary>
		/// Constructor with transaction parameter
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLImageRecordList(IDbTransaction trans): base(trans) {
		}

		/// <summary>
		/// Constructor with transaction parameter and alternate read method
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		/// <param name="alternateReadDBMethod"> Alternate Read Method</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLImageRecordList(IDbTransaction trans, ReadDBMethodDelegate alternateReadDBMethod): this(trans) {
			this.alternateReadDBMethod = alternateReadDBMethod;
		}

		/// <summary>
		/// Constructor declaration for serialization
		/// </summary>
		/// <param name="info"> SerializationInfo</param>
		/// <param name="context"> StreamingContext</param>
		public RLImageRecordList(SerializationInfo info, StreamingContext context): base(info, context) {
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
		protected override OSList<RCImageRecord> NewList() {
			return new RLImageRecordList();
		}


	} // RLImageRecordList

	/// <summary>
	/// RecordList type <code>RLProtectionRecordList</code> that represents a record list of
	///  <code>Protection</code>
	/// </summary>
	[Serializable()]
	public partial class RLProtectionRecordList: GenericRecordList<RCProtectionRecord>, IEnumerable, IEnumerator, ISerializable {
		public static void EnsureInitialized() {}

		protected override RCProtectionRecord GetElementDefaultValue() {
			return new RCProtectionRecord("");
		}

		public T[] ToArray<T>(Func<RCProtectionRecord, T> converter) {
			return ToArray(this, converter);
		}

		public static T[] ToArray<T>(RLProtectionRecordList recordlist, Func<RCProtectionRecord, T> converter) {
			return InnerToArray(recordlist, converter);
		}
		public static implicit operator RLProtectionRecordList(RCProtectionRecord[] array) {
			RLProtectionRecordList result = new RLProtectionRecordList();
			result.InnerFromArray(array);
			return result;
		}

		public static RLProtectionRecordList ToList<T>(T[] array, Func <T, RCProtectionRecord> converter) {
			RLProtectionRecordList result = new RLProtectionRecordList();
			result.InnerFromArray(array, converter);
			return result;
		}

		public static RLProtectionRecordList FromRestList<T>(RestList<T> restList, Func <T, RCProtectionRecord> converter) {
			RLProtectionRecordList result = new RLProtectionRecordList();
			result.InnerFromRestList(restList, converter);
			return result;
		}
		/// <summary>
		/// Default Constructor
		/// </summary>
		public RLProtectionRecordList(): base() {
		}

		/// <summary>
		/// Constructor with transaction parameter
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLProtectionRecordList(IDbTransaction trans): base(trans) {
		}

		/// <summary>
		/// Constructor with transaction parameter and alternate read method
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		/// <param name="alternateReadDBMethod"> Alternate Read Method</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLProtectionRecordList(IDbTransaction trans, ReadDBMethodDelegate alternateReadDBMethod): this(trans) {
			this.alternateReadDBMethod = alternateReadDBMethod;
		}

		/// <summary>
		/// Constructor declaration for serialization
		/// </summary>
		/// <param name="info"> SerializationInfo</param>
		/// <param name="context"> StreamingContext</param>
		public RLProtectionRecordList(SerializationInfo info, StreamingContext context): base(info, context) {
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
		protected override OSList<RCProtectionRecord> NewList() {
			return new RLProtectionRecordList();
		}


	} // RLProtectionRecordList

	/// <summary>
	/// RecordList type <code>RLValueRecordList</code> that represents a record list of <code>Value</code>
	/// </summary>
	[Serializable()]
	public partial class RLValueRecordList: GenericRecordList<RCValueRecord>, IEnumerable, IEnumerator, ISerializable {
		public static void EnsureInitialized() {}

		protected override RCValueRecord GetElementDefaultValue() {
			return new RCValueRecord("");
		}

		public T[] ToArray<T>(Func<RCValueRecord, T> converter) {
			return ToArray(this, converter);
		}

		public static T[] ToArray<T>(RLValueRecordList recordlist, Func<RCValueRecord, T> converter) {
			return InnerToArray(recordlist, converter);
		}
		public static implicit operator RLValueRecordList(RCValueRecord[] array) {
			RLValueRecordList result = new RLValueRecordList();
			result.InnerFromArray(array);
			return result;
		}

		public static RLValueRecordList ToList<T>(T[] array, Func <T, RCValueRecord> converter) {
			RLValueRecordList result = new RLValueRecordList();
			result.InnerFromArray(array, converter);
			return result;
		}

		public static RLValueRecordList FromRestList<T>(RestList<T> restList, Func <T, RCValueRecord> converter) {
			RLValueRecordList result = new RLValueRecordList();
			result.InnerFromRestList(restList, converter);
			return result;
		}
		/// <summary>
		/// Default Constructor
		/// </summary>
		public RLValueRecordList(): base() {
		}

		/// <summary>
		/// Constructor with transaction parameter
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLValueRecordList(IDbTransaction trans): base(trans) {
		}

		/// <summary>
		/// Constructor with transaction parameter and alternate read method
		/// </summary>
		/// <param name="trans"> IDbTransaction Parameter</param>
		/// <param name="alternateReadDBMethod"> Alternate Read Method</param>
		[Obsolete("Use the Default Constructor and set the Transaction afterwards.")]
		public RLValueRecordList(IDbTransaction trans, ReadDBMethodDelegate alternateReadDBMethod): this(trans) {
			this.alternateReadDBMethod = alternateReadDBMethod;
		}

		/// <summary>
		/// Constructor declaration for serialization
		/// </summary>
		/// <param name="info"> SerializationInfo</param>
		/// <param name="context"> StreamingContext</param>
		public RLValueRecordList(SerializationInfo info, StreamingContext context): base(info, context) {
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
		protected override OSList<RCValueRecord> NewList() {
			return new RLValueRecordList();
		}


	} // RLValueRecordList
}
