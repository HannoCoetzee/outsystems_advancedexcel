using System;
using System.Collections;
using System.Data;
using System.Runtime.Serialization;
using System.Reflection;
using System.Xml;
using OutSystems.ObjectKeys;
using OutSystems.RuntimeCommon;
using OutSystems.HubEdition.RuntimePlatform;
using OutSystems.HubEdition.RuntimePlatform.Db;
using OutSystems.Internal.Db;

namespace OutSystems.NssAdvanced_Excel {

	/// <summary>
	/// Structure <code>RCCellFormatRecord</code>
	/// </summary>
	[Serializable()]
	public partial struct RCCellFormatRecord: ISerializable, ITypedRecord<RCCellFormatRecord> {
		internal static readonly GlobalObjectKey IdCellFormat = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*_Kt8x60uod_oXugTPblTIQ");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("CellFormat")]
		public STCellFormatStructure ssSTCellFormat;


		public static implicit operator STCellFormatStructure(RCCellFormatRecord r) {
			return r.ssSTCellFormat;
		}

		public static implicit operator RCCellFormatRecord(STCellFormatStructure r) {
			RCCellFormatRecord res = new RCCellFormatRecord(null);
			res.ssSTCellFormat = r;
			return res;
		}

		public BitArray OptimizedAttributes;

		public RCCellFormatRecord(params string[] dummy) {
			OptimizedAttributes = null;
			ssSTCellFormat = new STCellFormatStructure(null);
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[1];
			all[0] = null;
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
					ssSTCellFormat.OptimizedAttributes = value[0];
				}
			}
			get {
				BitArray[] all = new BitArray[1];
				all[0] = null;
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssSTCellFormat.Read(r, ref index);
		}
		/// <summary>
		/// Read from database
		/// </summary>
		/// <param name="r"> Data reader</param>
		public void ReadDB(IDataReader r) {
			int index = 0;
			Read(r, ref index);
		}

		/// <summary>
		/// Read from record
		/// </summary>
		/// <param name="r"> Record</param>
		public void ReadIM(RCCellFormatRecord r) {
			this = r;
		}


		public static bool operator == (RCCellFormatRecord a, RCCellFormatRecord b) {
			if (a.ssSTCellFormat != b.ssSTCellFormat) return false;
			return true;
		}

		public static bool operator != (RCCellFormatRecord a, RCCellFormatRecord b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(RCCellFormatRecord)) return false;
			return (this == (RCCellFormatRecord) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssSTCellFormat.GetHashCode()
				;
			} catch {
				return base.GetHashCode();
			}
		}

		public void GetObjectData(SerializationInfo info, StreamingContext context) {
			Type objInfo = this.GetType();
			FieldInfo[] fields;
			fields = objInfo.GetFields(BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			for (int i = 0; i < fields.Length; i++)
			if (fields[i] .FieldType.IsSerializable)
			info.AddValue(fields[i] .Name, fields[i] .GetValue(this));
		}

		public RCCellFormatRecord(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssSTCellFormat = new STCellFormatStructure(null);
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssSTCellFormat", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssSTCellFormat' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssSTCellFormat = (STCellFormatStructure) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssSTCellFormat.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssSTCellFormat.InternalRecursiveSave();
		}


		public RCCellFormatRecord Duplicate() {
			RCCellFormatRecord t;
			t.ssSTCellFormat = (STCellFormatStructure) this.ssSTCellFormat.Duplicate();
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Record");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
			}
			if (detailLevel > 0) {
				ssSTCellFormat.ToXml(this, recordElem, "CellFormat", detailLevel - 1);
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "cellformat") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".CellFormat")) variable.Value = ssSTCellFormat; else variable.Optimized = true;
				variable.SetFieldName("cellformat");
			}
			if (variable.Found && tail != null) variable.EvaluateFields(this, head, tail);
		}

		public bool ChangedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public bool OptimizedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public object AttributeGet(GlobalObjectKey key) {
			if (key == IdCellFormat) {
				return ssSTCellFormat;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssSTCellFormat.FillFromOther((IRecord) other.AttributeGet(IdCellFormat));
		}
		public bool IsDefault() {
			RCCellFormatRecord defaultStruct = new RCCellFormatRecord(null);
			if (this.ssSTCellFormat != defaultStruct.ssSTCellFormat) return false;
			return true;
		}
	} // RCCellFormatRecord

	/// <summary>
	/// Structure <code>RCWorkbookRecord</code>
	/// </summary>
	[Serializable()]
	public partial struct RCWorkbookRecord: ISerializable, ITypedRecord<RCWorkbookRecord> {
		internal static readonly GlobalObjectKey IdWorkbook = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*cSfFgXnOmYAmiyqaFCX8Wg");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("Workbook")]
		public STWorkbookStructure ssSTWorkbook;


		public static implicit operator STWorkbookStructure(RCWorkbookRecord r) {
			return r.ssSTWorkbook;
		}

		public static implicit operator RCWorkbookRecord(STWorkbookStructure r) {
			RCWorkbookRecord res = new RCWorkbookRecord(null);
			res.ssSTWorkbook = r;
			return res;
		}

		public BitArray OptimizedAttributes;

		public RCWorkbookRecord(params string[] dummy) {
			OptimizedAttributes = null;
			ssSTWorkbook = new STWorkbookStructure(null);
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[1];
			all[0] = null;
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
					ssSTWorkbook.OptimizedAttributes = value[0];
				}
			}
			get {
				BitArray[] all = new BitArray[1];
				all[0] = null;
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssSTWorkbook.Read(r, ref index);
		}
		/// <summary>
		/// Read from database
		/// </summary>
		/// <param name="r"> Data reader</param>
		public void ReadDB(IDataReader r) {
			int index = 0;
			Read(r, ref index);
		}

		/// <summary>
		/// Read from record
		/// </summary>
		/// <param name="r"> Record</param>
		public void ReadIM(RCWorkbookRecord r) {
			this = r;
		}


		public static bool operator == (RCWorkbookRecord a, RCWorkbookRecord b) {
			if (a.ssSTWorkbook != b.ssSTWorkbook) return false;
			return true;
		}

		public static bool operator != (RCWorkbookRecord a, RCWorkbookRecord b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(RCWorkbookRecord)) return false;
			return (this == (RCWorkbookRecord) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssSTWorkbook.GetHashCode()
				;
			} catch {
				return base.GetHashCode();
			}
		}

		public void GetObjectData(SerializationInfo info, StreamingContext context) {
			Type objInfo = this.GetType();
			FieldInfo[] fields;
			fields = objInfo.GetFields(BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			for (int i = 0; i < fields.Length; i++)
			if (fields[i] .FieldType.IsSerializable)
			info.AddValue(fields[i] .Name, fields[i] .GetValue(this));
		}

		public RCWorkbookRecord(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssSTWorkbook = new STWorkbookStructure(null);
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssSTWorkbook", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssSTWorkbook' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssSTWorkbook = (STWorkbookStructure) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssSTWorkbook.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssSTWorkbook.InternalRecursiveSave();
		}


		public RCWorkbookRecord Duplicate() {
			RCWorkbookRecord t;
			t.ssSTWorkbook = (STWorkbookStructure) this.ssSTWorkbook.Duplicate();
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Record");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
			}
			if (detailLevel > 0) {
				ssSTWorkbook.ToXml(this, recordElem, "Workbook", detailLevel - 1);
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "workbook") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Workbook")) variable.Value = ssSTWorkbook; else variable.Optimized = true;
				variable.SetFieldName("workbook");
			}
			if (variable.Found && tail != null) variable.EvaluateFields(this, head, tail);
		}

		public bool ChangedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public bool OptimizedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public object AttributeGet(GlobalObjectKey key) {
			if (key == IdWorkbook) {
				return ssSTWorkbook;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssSTWorkbook.FillFromOther((IRecord) other.AttributeGet(IdWorkbook));
		}
		public bool IsDefault() {
			RCWorkbookRecord defaultStruct = new RCWorkbookRecord(null);
			if (this.ssSTWorkbook != defaultStruct.ssSTWorkbook) return false;
			return true;
		}
	} // RCWorkbookRecord

	/// <summary>
	/// Structure <code>RCWorksheetRecord</code>
	/// </summary>
	[Serializable()]
	public partial struct RCWorksheetRecord: ISerializable, ITypedRecord<RCWorksheetRecord> {
		internal static readonly GlobalObjectKey IdWorksheet = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*33h7wXL3Z32j+S7n4JJ2+g");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("Worksheet")]
		public STWorksheetStructure ssSTWorksheet;


		public static implicit operator STWorksheetStructure(RCWorksheetRecord r) {
			return r.ssSTWorksheet;
		}

		public static implicit operator RCWorksheetRecord(STWorksheetStructure r) {
			RCWorksheetRecord res = new RCWorksheetRecord(null);
			res.ssSTWorksheet = r;
			return res;
		}

		public BitArray OptimizedAttributes;

		public RCWorksheetRecord(params string[] dummy) {
			OptimizedAttributes = null;
			ssSTWorksheet = new STWorksheetStructure(null);
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[1];
			all[0] = null;
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
					ssSTWorksheet.OptimizedAttributes = value[0];
				}
			}
			get {
				BitArray[] all = new BitArray[1];
				all[0] = null;
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssSTWorksheet.Read(r, ref index);
		}
		/// <summary>
		/// Read from database
		/// </summary>
		/// <param name="r"> Data reader</param>
		public void ReadDB(IDataReader r) {
			int index = 0;
			Read(r, ref index);
		}

		/// <summary>
		/// Read from record
		/// </summary>
		/// <param name="r"> Record</param>
		public void ReadIM(RCWorksheetRecord r) {
			this = r;
		}


		public static bool operator == (RCWorksheetRecord a, RCWorksheetRecord b) {
			if (a.ssSTWorksheet != b.ssSTWorksheet) return false;
			return true;
		}

		public static bool operator != (RCWorksheetRecord a, RCWorksheetRecord b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(RCWorksheetRecord)) return false;
			return (this == (RCWorksheetRecord) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssSTWorksheet.GetHashCode()
				;
			} catch {
				return base.GetHashCode();
			}
		}

		public void GetObjectData(SerializationInfo info, StreamingContext context) {
			Type objInfo = this.GetType();
			FieldInfo[] fields;
			fields = objInfo.GetFields(BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			for (int i = 0; i < fields.Length; i++)
			if (fields[i] .FieldType.IsSerializable)
			info.AddValue(fields[i] .Name, fields[i] .GetValue(this));
		}

		public RCWorksheetRecord(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssSTWorksheet = new STWorksheetStructure(null);
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssSTWorksheet", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssSTWorksheet' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssSTWorksheet = (STWorksheetStructure) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssSTWorksheet.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssSTWorksheet.InternalRecursiveSave();
		}


		public RCWorksheetRecord Duplicate() {
			RCWorksheetRecord t;
			t.ssSTWorksheet = (STWorksheetStructure) this.ssSTWorksheet.Duplicate();
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Record");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
			}
			if (detailLevel > 0) {
				ssSTWorksheet.ToXml(this, recordElem, "Worksheet", detailLevel - 1);
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "worksheet") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Worksheet")) variable.Value = ssSTWorksheet; else variable.Optimized = true;
				variable.SetFieldName("worksheet");
			}
			if (variable.Found && tail != null) variable.EvaluateFields(this, head, tail);
		}

		public bool ChangedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public bool OptimizedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public object AttributeGet(GlobalObjectKey key) {
			if (key == IdWorksheet) {
				return ssSTWorksheet;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssSTWorksheet.FillFromOther((IRecord) other.AttributeGet(IdWorksheet));
		}
		public bool IsDefault() {
			RCWorksheetRecord defaultStruct = new RCWorksheetRecord(null);
			if (this.ssSTWorksheet != defaultStruct.ssSTWorksheet) return false;
			return true;
		}
	} // RCWorksheetRecord

	/// <summary>
	/// Structure <code>RCColorRecord</code>
	/// </summary>
	[Serializable()]
	public partial struct RCColorRecord: ISerializable, ITypedRecord<RCColorRecord> {
		internal static readonly GlobalObjectKey IdColor = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*egnk0hJuQI_VWFqkbK8pLw");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("Color")]
		public STColorStructure ssSTColor;


		public static implicit operator STColorStructure(RCColorRecord r) {
			return r.ssSTColor;
		}

		public static implicit operator RCColorRecord(STColorStructure r) {
			RCColorRecord res = new RCColorRecord(null);
			res.ssSTColor = r;
			return res;
		}

		public BitArray OptimizedAttributes;

		public RCColorRecord(params string[] dummy) {
			OptimizedAttributes = null;
			ssSTColor = new STColorStructure(null);
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[1];
			all[0] = null;
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
					ssSTColor.OptimizedAttributes = value[0];
				}
			}
			get {
				BitArray[] all = new BitArray[1];
				all[0] = null;
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssSTColor.Read(r, ref index);
		}
		/// <summary>
		/// Read from database
		/// </summary>
		/// <param name="r"> Data reader</param>
		public void ReadDB(IDataReader r) {
			int index = 0;
			Read(r, ref index);
		}

		/// <summary>
		/// Read from record
		/// </summary>
		/// <param name="r"> Record</param>
		public void ReadIM(RCColorRecord r) {
			this = r;
		}


		public static bool operator == (RCColorRecord a, RCColorRecord b) {
			if (a.ssSTColor != b.ssSTColor) return false;
			return true;
		}

		public static bool operator != (RCColorRecord a, RCColorRecord b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(RCColorRecord)) return false;
			return (this == (RCColorRecord) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssSTColor.GetHashCode()
				;
			} catch {
				return base.GetHashCode();
			}
		}

		public void GetObjectData(SerializationInfo info, StreamingContext context) {
			Type objInfo = this.GetType();
			FieldInfo[] fields;
			fields = objInfo.GetFields(BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			for (int i = 0; i < fields.Length; i++)
			if (fields[i] .FieldType.IsSerializable)
			info.AddValue(fields[i] .Name, fields[i] .GetValue(this));
		}

		public RCColorRecord(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssSTColor = new STColorStructure(null);
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssSTColor", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssSTColor' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssSTColor = (STColorStructure) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssSTColor.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssSTColor.InternalRecursiveSave();
		}


		public RCColorRecord Duplicate() {
			RCColorRecord t;
			t.ssSTColor = (STColorStructure) this.ssSTColor.Duplicate();
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Record");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
			}
			if (detailLevel > 0) {
				ssSTColor.ToXml(this, recordElem, "Color", detailLevel - 1);
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "color") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Color")) variable.Value = ssSTColor; else variable.Optimized = true;
				variable.SetFieldName("color");
			}
			if (variable.Found && tail != null) variable.EvaluateFields(this, head, tail);
		}

		public bool ChangedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public bool OptimizedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public object AttributeGet(GlobalObjectKey key) {
			if (key == IdColor) {
				return ssSTColor;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssSTColor.FillFromOther((IRecord) other.AttributeGet(IdColor));
		}
		public bool IsDefault() {
			RCColorRecord defaultStruct = new RCColorRecord(null);
			if (this.ssSTColor != defaultStruct.ssSTColor) return false;
			return true;
		}
	} // RCColorRecord

	/// <summary>
	/// Structure <code>RCDimensionRecord</code>
	/// </summary>
	[Serializable()]
	public partial struct RCDimensionRecord: ISerializable, ITypedRecord<RCDimensionRecord> {
		internal static readonly GlobalObjectKey IdDimension = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*GgUWI0Z2l9Rs7FF+CSAoGQ");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("Dimension")]
		public STDimensionStructure ssSTDimension;


		public static implicit operator STDimensionStructure(RCDimensionRecord r) {
			return r.ssSTDimension;
		}

		public static implicit operator RCDimensionRecord(STDimensionStructure r) {
			RCDimensionRecord res = new RCDimensionRecord(null);
			res.ssSTDimension = r;
			return res;
		}

		public BitArray OptimizedAttributes;

		public RCDimensionRecord(params string[] dummy) {
			OptimizedAttributes = null;
			ssSTDimension = new STDimensionStructure(null);
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[1];
			all[0] = null;
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
					ssSTDimension.OptimizedAttributes = value[0];
				}
			}
			get {
				BitArray[] all = new BitArray[1];
				all[0] = null;
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssSTDimension.Read(r, ref index);
		}
		/// <summary>
		/// Read from database
		/// </summary>
		/// <param name="r"> Data reader</param>
		public void ReadDB(IDataReader r) {
			int index = 0;
			Read(r, ref index);
		}

		/// <summary>
		/// Read from record
		/// </summary>
		/// <param name="r"> Record</param>
		public void ReadIM(RCDimensionRecord r) {
			this = r;
		}


		public static bool operator == (RCDimensionRecord a, RCDimensionRecord b) {
			if (a.ssSTDimension != b.ssSTDimension) return false;
			return true;
		}

		public static bool operator != (RCDimensionRecord a, RCDimensionRecord b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(RCDimensionRecord)) return false;
			return (this == (RCDimensionRecord) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssSTDimension.GetHashCode()
				;
			} catch {
				return base.GetHashCode();
			}
		}

		public void GetObjectData(SerializationInfo info, StreamingContext context) {
			Type objInfo = this.GetType();
			FieldInfo[] fields;
			fields = objInfo.GetFields(BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			for (int i = 0; i < fields.Length; i++)
			if (fields[i] .FieldType.IsSerializable)
			info.AddValue(fields[i] .Name, fields[i] .GetValue(this));
		}

		public RCDimensionRecord(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssSTDimension = new STDimensionStructure(null);
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssSTDimension", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssSTDimension' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssSTDimension = (STDimensionStructure) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssSTDimension.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssSTDimension.InternalRecursiveSave();
		}


		public RCDimensionRecord Duplicate() {
			RCDimensionRecord t;
			t.ssSTDimension = (STDimensionStructure) this.ssSTDimension.Duplicate();
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Record");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
			}
			if (detailLevel > 0) {
				ssSTDimension.ToXml(this, recordElem, "Dimension", detailLevel - 1);
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "dimension") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Dimension")) variable.Value = ssSTDimension; else variable.Optimized = true;
				variable.SetFieldName("dimension");
			}
			if (variable.Found && tail != null) variable.EvaluateFields(this, head, tail);
		}

		public bool ChangedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public bool OptimizedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public object AttributeGet(GlobalObjectKey key) {
			if (key == IdDimension) {
				return ssSTDimension;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssSTDimension.FillFromOther((IRecord) other.AttributeGet(IdDimension));
		}
		public bool IsDefault() {
			RCDimensionRecord defaultStruct = new RCDimensionRecord(null);
			if (this.ssSTDimension != defaultStruct.ssSTDimension) return false;
			return true;
		}
	} // RCDimensionRecord

	/// <summary>
	/// Structure <code>RCAddressRecord</code>
	/// </summary>
	[Serializable()]
	public partial struct RCAddressRecord: ISerializable, ITypedRecord<RCAddressRecord> {
		internal static readonly GlobalObjectKey IdAddress = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*sakV9VS1OspLz+KlcvXSag");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("Address")]
		public STAddressStructure ssSTAddress;


		public static implicit operator STAddressStructure(RCAddressRecord r) {
			return r.ssSTAddress;
		}

		public static implicit operator RCAddressRecord(STAddressStructure r) {
			RCAddressRecord res = new RCAddressRecord(null);
			res.ssSTAddress = r;
			return res;
		}

		public BitArray OptimizedAttributes;

		public RCAddressRecord(params string[] dummy) {
			OptimizedAttributes = null;
			ssSTAddress = new STAddressStructure(null);
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[1];
			all[0] = null;
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
					ssSTAddress.OptimizedAttributes = value[0];
				}
			}
			get {
				BitArray[] all = new BitArray[1];
				all[0] = null;
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssSTAddress.Read(r, ref index);
		}
		/// <summary>
		/// Read from database
		/// </summary>
		/// <param name="r"> Data reader</param>
		public void ReadDB(IDataReader r) {
			int index = 0;
			Read(r, ref index);
		}

		/// <summary>
		/// Read from record
		/// </summary>
		/// <param name="r"> Record</param>
		public void ReadIM(RCAddressRecord r) {
			this = r;
		}


		public static bool operator == (RCAddressRecord a, RCAddressRecord b) {
			if (a.ssSTAddress != b.ssSTAddress) return false;
			return true;
		}

		public static bool operator != (RCAddressRecord a, RCAddressRecord b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(RCAddressRecord)) return false;
			return (this == (RCAddressRecord) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssSTAddress.GetHashCode()
				;
			} catch {
				return base.GetHashCode();
			}
		}

		public void GetObjectData(SerializationInfo info, StreamingContext context) {
			Type objInfo = this.GetType();
			FieldInfo[] fields;
			fields = objInfo.GetFields(BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			for (int i = 0; i < fields.Length; i++)
			if (fields[i] .FieldType.IsSerializable)
			info.AddValue(fields[i] .Name, fields[i] .GetValue(this));
		}

		public RCAddressRecord(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssSTAddress = new STAddressStructure(null);
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssSTAddress", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssSTAddress' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssSTAddress = (STAddressStructure) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssSTAddress.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssSTAddress.InternalRecursiveSave();
		}


		public RCAddressRecord Duplicate() {
			RCAddressRecord t;
			t.ssSTAddress = (STAddressStructure) this.ssSTAddress.Duplicate();
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Record");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
			}
			if (detailLevel > 0) {
				ssSTAddress.ToXml(this, recordElem, "Address", detailLevel - 1);
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "address") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Address")) variable.Value = ssSTAddress; else variable.Optimized = true;
				variable.SetFieldName("address");
			}
			if (variable.Found && tail != null) variable.EvaluateFields(this, head, tail);
		}

		public bool ChangedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public bool OptimizedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public object AttributeGet(GlobalObjectKey key) {
			if (key == IdAddress) {
				return ssSTAddress;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssSTAddress.FillFromOther((IRecord) other.AttributeGet(IdAddress));
		}
		public bool IsDefault() {
			RCAddressRecord defaultStruct = new RCAddressRecord(null);
			if (this.ssSTAddress != defaultStruct.ssSTAddress) return false;
			return true;
		}
	} // RCAddressRecord

	/// <summary>
	/// Structure <code>RCRangeRecord</code>
	/// </summary>
	[Serializable()]
	public partial struct RCRangeRecord: ISerializable, ITypedRecord<RCRangeRecord> {
		internal static readonly GlobalObjectKey IdRange = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*fkdXsiMILofCapOVw0TaOg");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("Range")]
		public STRangeStructure ssSTRange;


		public static implicit operator STRangeStructure(RCRangeRecord r) {
			return r.ssSTRange;
		}

		public static implicit operator RCRangeRecord(STRangeStructure r) {
			RCRangeRecord res = new RCRangeRecord(null);
			res.ssSTRange = r;
			return res;
		}

		public BitArray OptimizedAttributes;

		public RCRangeRecord(params string[] dummy) {
			OptimizedAttributes = null;
			ssSTRange = new STRangeStructure(null);
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[1];
			all[0] = null;
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
					ssSTRange.OptimizedAttributes = value[0];
				}
			}
			get {
				BitArray[] all = new BitArray[1];
				all[0] = null;
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssSTRange.Read(r, ref index);
		}
		/// <summary>
		/// Read from database
		/// </summary>
		/// <param name="r"> Data reader</param>
		public void ReadDB(IDataReader r) {
			int index = 0;
			Read(r, ref index);
		}

		/// <summary>
		/// Read from record
		/// </summary>
		/// <param name="r"> Record</param>
		public void ReadIM(RCRangeRecord r) {
			this = r;
		}


		public static bool operator == (RCRangeRecord a, RCRangeRecord b) {
			if (a.ssSTRange != b.ssSTRange) return false;
			return true;
		}

		public static bool operator != (RCRangeRecord a, RCRangeRecord b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(RCRangeRecord)) return false;
			return (this == (RCRangeRecord) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssSTRange.GetHashCode()
				;
			} catch {
				return base.GetHashCode();
			}
		}

		public void GetObjectData(SerializationInfo info, StreamingContext context) {
			Type objInfo = this.GetType();
			FieldInfo[] fields;
			fields = objInfo.GetFields(BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			for (int i = 0; i < fields.Length; i++)
			if (fields[i] .FieldType.IsSerializable)
			info.AddValue(fields[i] .Name, fields[i] .GetValue(this));
		}

		public RCRangeRecord(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssSTRange = new STRangeStructure(null);
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssSTRange", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssSTRange' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssSTRange = (STRangeStructure) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssSTRange.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssSTRange.InternalRecursiveSave();
		}


		public RCRangeRecord Duplicate() {
			RCRangeRecord t;
			t.ssSTRange = (STRangeStructure) this.ssSTRange.Duplicate();
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Record");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
			}
			if (detailLevel > 0) {
				ssSTRange.ToXml(this, recordElem, "Range", detailLevel - 1);
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "range") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Range")) variable.Value = ssSTRange; else variable.Optimized = true;
				variable.SetFieldName("range");
			}
			if (variable.Found && tail != null) variable.EvaluateFields(this, head, tail);
		}

		public bool ChangedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public bool OptimizedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public object AttributeGet(GlobalObjectKey key) {
			if (key == IdRange) {
				return ssSTRange;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssSTRange.FillFromOther((IRecord) other.AttributeGet(IdRange));
		}
		public bool IsDefault() {
			RCRangeRecord defaultStruct = new RCRangeRecord(null);
			if (this.ssSTRange != defaultStruct.ssSTRange) return false;
			return true;
		}
	} // RCRangeRecord

	/// <summary>
	/// Structure <code>RCDataSeriesRecord</code>
	/// </summary>
	[Serializable()]
	public partial struct RCDataSeriesRecord: ISerializable, ITypedRecord<RCDataSeriesRecord> {
		internal static readonly GlobalObjectKey IdDataSeries = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*8aE7EofndyiSFdNxWsJcng");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("DataSeries")]
		public STDataSeriesStructure ssSTDataSeries;


		public static implicit operator STDataSeriesStructure(RCDataSeriesRecord r) {
			return r.ssSTDataSeries;
		}

		public static implicit operator RCDataSeriesRecord(STDataSeriesStructure r) {
			RCDataSeriesRecord res = new RCDataSeriesRecord(null);
			res.ssSTDataSeries = r;
			return res;
		}

		public BitArray OptimizedAttributes;

		public RCDataSeriesRecord(params string[] dummy) {
			OptimizedAttributes = null;
			ssSTDataSeries = new STDataSeriesStructure(null);
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[1];
			all[0] = null;
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
					ssSTDataSeries.OptimizedAttributes = value[0];
				}
			}
			get {
				BitArray[] all = new BitArray[1];
				all[0] = null;
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssSTDataSeries.Read(r, ref index);
		}
		/// <summary>
		/// Read from database
		/// </summary>
		/// <param name="r"> Data reader</param>
		public void ReadDB(IDataReader r) {
			int index = 0;
			Read(r, ref index);
		}

		/// <summary>
		/// Read from record
		/// </summary>
		/// <param name="r"> Record</param>
		public void ReadIM(RCDataSeriesRecord r) {
			this = r;
		}


		public static bool operator == (RCDataSeriesRecord a, RCDataSeriesRecord b) {
			if (a.ssSTDataSeries != b.ssSTDataSeries) return false;
			return true;
		}

		public static bool operator != (RCDataSeriesRecord a, RCDataSeriesRecord b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(RCDataSeriesRecord)) return false;
			return (this == (RCDataSeriesRecord) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssSTDataSeries.GetHashCode()
				;
			} catch {
				return base.GetHashCode();
			}
		}

		public void GetObjectData(SerializationInfo info, StreamingContext context) {
			Type objInfo = this.GetType();
			FieldInfo[] fields;
			fields = objInfo.GetFields(BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			for (int i = 0; i < fields.Length; i++)
			if (fields[i] .FieldType.IsSerializable)
			info.AddValue(fields[i] .Name, fields[i] .GetValue(this));
		}

		public RCDataSeriesRecord(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssSTDataSeries = new STDataSeriesStructure(null);
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssSTDataSeries", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssSTDataSeries' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssSTDataSeries = (STDataSeriesStructure) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssSTDataSeries.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssSTDataSeries.InternalRecursiveSave();
		}


		public RCDataSeriesRecord Duplicate() {
			RCDataSeriesRecord t;
			t.ssSTDataSeries = (STDataSeriesStructure) this.ssSTDataSeries.Duplicate();
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Record");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
			}
			if (detailLevel > 0) {
				ssSTDataSeries.ToXml(this, recordElem, "DataSeries", detailLevel - 1);
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "dataseries") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".DataSeries")) variable.Value = ssSTDataSeries; else variable.Optimized = true;
				variable.SetFieldName("dataseries");
			}
			if (variable.Found && tail != null) variable.EvaluateFields(this, head, tail);
		}

		public bool ChangedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public bool OptimizedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public object AttributeGet(GlobalObjectKey key) {
			if (key == IdDataSeries) {
				return ssSTDataSeries;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssSTDataSeries.FillFromOther((IRecord) other.AttributeGet(IdDataSeries));
		}
		public bool IsDefault() {
			RCDataSeriesRecord defaultStruct = new RCDataSeriesRecord(null);
			if (this.ssSTDataSeries != defaultStruct.ssSTDataSeries) return false;
			return true;
		}
	} // RCDataSeriesRecord

	/// <summary>
	/// Structure <code>RCConditionalFormatItemRecord</code>
	/// </summary>
	[Serializable()]
	public partial struct RCConditionalFormatItemRecord: ISerializable, ITypedRecord<RCConditionalFormatItemRecord> {
		internal static readonly GlobalObjectKey IdConditionalFormatItem = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*m7UhTXLoFcESauNFfSz_9Q");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("ConditionalFormatItem")]
		public STConditionalFormatItemStructure ssSTConditionalFormatItem;


		public static implicit operator STConditionalFormatItemStructure(RCConditionalFormatItemRecord r) {
			return r.ssSTConditionalFormatItem;
		}

		public static implicit operator RCConditionalFormatItemRecord(STConditionalFormatItemStructure r) {
			RCConditionalFormatItemRecord res = new RCConditionalFormatItemRecord(null);
			res.ssSTConditionalFormatItem = r;
			return res;
		}

		public BitArray OptimizedAttributes;

		public RCConditionalFormatItemRecord(params string[] dummy) {
			OptimizedAttributes = null;
			ssSTConditionalFormatItem = new STConditionalFormatItemStructure(null);
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[1];
			all[0] = null;
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
					ssSTConditionalFormatItem.OptimizedAttributes = value[0];
				}
			}
			get {
				BitArray[] all = new BitArray[1];
				all[0] = null;
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssSTConditionalFormatItem.Read(r, ref index);
		}
		/// <summary>
		/// Read from database
		/// </summary>
		/// <param name="r"> Data reader</param>
		public void ReadDB(IDataReader r) {
			int index = 0;
			Read(r, ref index);
		}

		/// <summary>
		/// Read from record
		/// </summary>
		/// <param name="r"> Record</param>
		public void ReadIM(RCConditionalFormatItemRecord r) {
			this = r;
		}


		public static bool operator == (RCConditionalFormatItemRecord a, RCConditionalFormatItemRecord b) {
			if (a.ssSTConditionalFormatItem != b.ssSTConditionalFormatItem) return false;
			return true;
		}

		public static bool operator != (RCConditionalFormatItemRecord a, RCConditionalFormatItemRecord b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(RCConditionalFormatItemRecord)) return false;
			return (this == (RCConditionalFormatItemRecord) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssSTConditionalFormatItem.GetHashCode()
				;
			} catch {
				return base.GetHashCode();
			}
		}

		public void GetObjectData(SerializationInfo info, StreamingContext context) {
			Type objInfo = this.GetType();
			FieldInfo[] fields;
			fields = objInfo.GetFields(BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			for (int i = 0; i < fields.Length; i++)
			if (fields[i] .FieldType.IsSerializable)
			info.AddValue(fields[i] .Name, fields[i] .GetValue(this));
		}

		public RCConditionalFormatItemRecord(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssSTConditionalFormatItem = new STConditionalFormatItemStructure(null);
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssSTConditionalFormatItem", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssSTConditionalFormatItem' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssSTConditionalFormatItem = (STConditionalFormatItemStructure) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssSTConditionalFormatItem.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssSTConditionalFormatItem.InternalRecursiveSave();
		}


		public RCConditionalFormatItemRecord Duplicate() {
			RCConditionalFormatItemRecord t;
			t.ssSTConditionalFormatItem = (STConditionalFormatItemStructure) this.ssSTConditionalFormatItem.Duplicate();
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Record");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
			}
			if (detailLevel > 0) {
				ssSTConditionalFormatItem.ToXml(this, recordElem, "ConditionalFormatItem", detailLevel - 1);
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "conditionalformatitem") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".ConditionalFormatItem")) variable.Value = ssSTConditionalFormatItem; else variable.Optimized = true;
				variable.SetFieldName("conditionalformatitem");
			}
			if (variable.Found && tail != null) variable.EvaluateFields(this, head, tail);
		}

		public bool ChangedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public bool OptimizedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public object AttributeGet(GlobalObjectKey key) {
			if (key == IdConditionalFormatItem) {
				return ssSTConditionalFormatItem;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssSTConditionalFormatItem.FillFromOther((IRecord) other.AttributeGet(IdConditionalFormatItem));
		}
		public bool IsDefault() {
			RCConditionalFormatItemRecord defaultStruct = new RCConditionalFormatItemRecord(null);
			if (this.ssSTConditionalFormatItem != defaultStruct.ssSTConditionalFormatItem) return false;
			return true;
		}
	} // RCConditionalFormatItemRecord

	/// <summary>
	/// Structure <code>RCConditionalFormatStyleRecord</code>
	/// </summary>
	[Serializable()]
	public partial struct RCConditionalFormatStyleRecord: ISerializable, ITypedRecord<RCConditionalFormatStyleRecord> {
		internal static readonly GlobalObjectKey IdConditionalFormatStyle = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*qsmbMO5tGXSuEne6oWL2FA");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("ConditionalFormatStyle")]
		public STConditionalFormatStyleStructure ssSTConditionalFormatStyle;


		public static implicit operator STConditionalFormatStyleStructure(RCConditionalFormatStyleRecord r) {
			return r.ssSTConditionalFormatStyle;
		}

		public static implicit operator RCConditionalFormatStyleRecord(STConditionalFormatStyleStructure r) {
			RCConditionalFormatStyleRecord res = new RCConditionalFormatStyleRecord(null);
			res.ssSTConditionalFormatStyle = r;
			return res;
		}

		public BitArray OptimizedAttributes;

		public RCConditionalFormatStyleRecord(params string[] dummy) {
			OptimizedAttributes = null;
			ssSTConditionalFormatStyle = new STConditionalFormatStyleStructure(null);
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[1];
			all[0] = null;
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
					ssSTConditionalFormatStyle.OptimizedAttributes = value[0];
				}
			}
			get {
				BitArray[] all = new BitArray[1];
				all[0] = null;
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssSTConditionalFormatStyle.Read(r, ref index);
		}
		/// <summary>
		/// Read from database
		/// </summary>
		/// <param name="r"> Data reader</param>
		public void ReadDB(IDataReader r) {
			int index = 0;
			Read(r, ref index);
		}

		/// <summary>
		/// Read from record
		/// </summary>
		/// <param name="r"> Record</param>
		public void ReadIM(RCConditionalFormatStyleRecord r) {
			this = r;
		}


		public static bool operator == (RCConditionalFormatStyleRecord a, RCConditionalFormatStyleRecord b) {
			if (a.ssSTConditionalFormatStyle != b.ssSTConditionalFormatStyle) return false;
			return true;
		}

		public static bool operator != (RCConditionalFormatStyleRecord a, RCConditionalFormatStyleRecord b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(RCConditionalFormatStyleRecord)) return false;
			return (this == (RCConditionalFormatStyleRecord) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssSTConditionalFormatStyle.GetHashCode()
				;
			} catch {
				return base.GetHashCode();
			}
		}

		public void GetObjectData(SerializationInfo info, StreamingContext context) {
			Type objInfo = this.GetType();
			FieldInfo[] fields;
			fields = objInfo.GetFields(BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			for (int i = 0; i < fields.Length; i++)
			if (fields[i] .FieldType.IsSerializable)
			info.AddValue(fields[i] .Name, fields[i] .GetValue(this));
		}

		public RCConditionalFormatStyleRecord(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssSTConditionalFormatStyle = new STConditionalFormatStyleStructure(null);
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssSTConditionalFormatStyle", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssSTConditionalFormatStyle' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssSTConditionalFormatStyle = (STConditionalFormatStyleStructure) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssSTConditionalFormatStyle.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssSTConditionalFormatStyle.InternalRecursiveSave();
		}


		public RCConditionalFormatStyleRecord Duplicate() {
			RCConditionalFormatStyleRecord t;
			t.ssSTConditionalFormatStyle = (STConditionalFormatStyleStructure) this.ssSTConditionalFormatStyle.Duplicate();
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Record");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
			}
			if (detailLevel > 0) {
				ssSTConditionalFormatStyle.ToXml(this, recordElem, "ConditionalFormatStyle", detailLevel - 1);
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "conditionalformatstyle") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".ConditionalFormatStyle")) variable.Value = ssSTConditionalFormatStyle; else variable.Optimized = true;
				variable.SetFieldName("conditionalformatstyle");
			}
			if (variable.Found && tail != null) variable.EvaluateFields(this, head, tail);
		}

		public bool ChangedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public bool OptimizedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public object AttributeGet(GlobalObjectKey key) {
			if (key == IdConditionalFormatStyle) {
				return ssSTConditionalFormatStyle;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssSTConditionalFormatStyle.FillFromOther((IRecord) other.AttributeGet(IdConditionalFormatStyle));
		}
		public bool IsDefault() {
			RCConditionalFormatStyleRecord defaultStruct = new RCConditionalFormatStyleRecord(null);
			if (this.ssSTConditionalFormatStyle != defaultStruct.ssSTConditionalFormatStyle) return false;
			return true;
		}
	} // RCConditionalFormatStyleRecord

	/// <summary>
	/// Structure <code>RCBorderStyleRecord</code>
	/// </summary>
	[Serializable()]
	public partial struct RCBorderStyleRecord: ISerializable, ITypedRecord<RCBorderStyleRecord> {
		internal static readonly GlobalObjectKey IdBorderStyle = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*+qrCrqrouya0sitZYLHvTA");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("BorderStyle")]
		public STBorderStyleStructure ssSTBorderStyle;


		public static implicit operator STBorderStyleStructure(RCBorderStyleRecord r) {
			return r.ssSTBorderStyle;
		}

		public static implicit operator RCBorderStyleRecord(STBorderStyleStructure r) {
			RCBorderStyleRecord res = new RCBorderStyleRecord(null);
			res.ssSTBorderStyle = r;
			return res;
		}

		public BitArray OptimizedAttributes;

		public RCBorderStyleRecord(params string[] dummy) {
			OptimizedAttributes = null;
			ssSTBorderStyle = new STBorderStyleStructure(null);
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[1];
			all[0] = null;
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
					ssSTBorderStyle.OptimizedAttributes = value[0];
				}
			}
			get {
				BitArray[] all = new BitArray[1];
				all[0] = null;
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssSTBorderStyle.Read(r, ref index);
		}
		/// <summary>
		/// Read from database
		/// </summary>
		/// <param name="r"> Data reader</param>
		public void ReadDB(IDataReader r) {
			int index = 0;
			Read(r, ref index);
		}

		/// <summary>
		/// Read from record
		/// </summary>
		/// <param name="r"> Record</param>
		public void ReadIM(RCBorderStyleRecord r) {
			this = r;
		}


		public static bool operator == (RCBorderStyleRecord a, RCBorderStyleRecord b) {
			if (a.ssSTBorderStyle != b.ssSTBorderStyle) return false;
			return true;
		}

		public static bool operator != (RCBorderStyleRecord a, RCBorderStyleRecord b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(RCBorderStyleRecord)) return false;
			return (this == (RCBorderStyleRecord) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssSTBorderStyle.GetHashCode()
				;
			} catch {
				return base.GetHashCode();
			}
		}

		public void GetObjectData(SerializationInfo info, StreamingContext context) {
			Type objInfo = this.GetType();
			FieldInfo[] fields;
			fields = objInfo.GetFields(BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			for (int i = 0; i < fields.Length; i++)
			if (fields[i] .FieldType.IsSerializable)
			info.AddValue(fields[i] .Name, fields[i] .GetValue(this));
		}

		public RCBorderStyleRecord(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssSTBorderStyle = new STBorderStyleStructure(null);
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssSTBorderStyle", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssSTBorderStyle' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssSTBorderStyle = (STBorderStyleStructure) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssSTBorderStyle.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssSTBorderStyle.InternalRecursiveSave();
		}


		public RCBorderStyleRecord Duplicate() {
			RCBorderStyleRecord t;
			t.ssSTBorderStyle = (STBorderStyleStructure) this.ssSTBorderStyle.Duplicate();
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Record");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
			}
			if (detailLevel > 0) {
				ssSTBorderStyle.ToXml(this, recordElem, "BorderStyle", detailLevel - 1);
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "borderstyle") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".BorderStyle")) variable.Value = ssSTBorderStyle; else variable.Optimized = true;
				variable.SetFieldName("borderstyle");
			}
			if (variable.Found && tail != null) variable.EvaluateFields(this, head, tail);
		}

		public bool ChangedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public bool OptimizedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public object AttributeGet(GlobalObjectKey key) {
			if (key == IdBorderStyle) {
				return ssSTBorderStyle;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssSTBorderStyle.FillFromOther((IRecord) other.AttributeGet(IdBorderStyle));
		}
		public bool IsDefault() {
			RCBorderStyleRecord defaultStruct = new RCBorderStyleRecord(null);
			if (this.ssSTBorderStyle != defaultStruct.ssSTBorderStyle) return false;
			return true;
		}
	} // RCBorderStyleRecord

	/// <summary>
	/// Structure <code>RCFillStyleRecord</code>
	/// </summary>
	[Serializable()]
	public partial struct RCFillStyleRecord: ISerializable, ITypedRecord<RCFillStyleRecord> {
		internal static readonly GlobalObjectKey IdFillStyle = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*_XAc2AFE5eSXijSex78YQw");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("FillStyle")]
		public STFillStyleStructure ssSTFillStyle;


		public static implicit operator STFillStyleStructure(RCFillStyleRecord r) {
			return r.ssSTFillStyle;
		}

		public static implicit operator RCFillStyleRecord(STFillStyleStructure r) {
			RCFillStyleRecord res = new RCFillStyleRecord(null);
			res.ssSTFillStyle = r;
			return res;
		}

		public BitArray OptimizedAttributes;

		public RCFillStyleRecord(params string[] dummy) {
			OptimizedAttributes = null;
			ssSTFillStyle = new STFillStyleStructure(null);
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[1];
			all[0] = null;
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
					ssSTFillStyle.OptimizedAttributes = value[0];
				}
			}
			get {
				BitArray[] all = new BitArray[1];
				all[0] = null;
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssSTFillStyle.Read(r, ref index);
		}
		/// <summary>
		/// Read from database
		/// </summary>
		/// <param name="r"> Data reader</param>
		public void ReadDB(IDataReader r) {
			int index = 0;
			Read(r, ref index);
		}

		/// <summary>
		/// Read from record
		/// </summary>
		/// <param name="r"> Record</param>
		public void ReadIM(RCFillStyleRecord r) {
			this = r;
		}


		public static bool operator == (RCFillStyleRecord a, RCFillStyleRecord b) {
			if (a.ssSTFillStyle != b.ssSTFillStyle) return false;
			return true;
		}

		public static bool operator != (RCFillStyleRecord a, RCFillStyleRecord b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(RCFillStyleRecord)) return false;
			return (this == (RCFillStyleRecord) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssSTFillStyle.GetHashCode()
				;
			} catch {
				return base.GetHashCode();
			}
		}

		public void GetObjectData(SerializationInfo info, StreamingContext context) {
			Type objInfo = this.GetType();
			FieldInfo[] fields;
			fields = objInfo.GetFields(BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			for (int i = 0; i < fields.Length; i++)
			if (fields[i] .FieldType.IsSerializable)
			info.AddValue(fields[i] .Name, fields[i] .GetValue(this));
		}

		public RCFillStyleRecord(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssSTFillStyle = new STFillStyleStructure(null);
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssSTFillStyle", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssSTFillStyle' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssSTFillStyle = (STFillStyleStructure) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssSTFillStyle.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssSTFillStyle.InternalRecursiveSave();
		}


		public RCFillStyleRecord Duplicate() {
			RCFillStyleRecord t;
			t.ssSTFillStyle = (STFillStyleStructure) this.ssSTFillStyle.Duplicate();
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Record");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
			}
			if (detailLevel > 0) {
				ssSTFillStyle.ToXml(this, recordElem, "FillStyle", detailLevel - 1);
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "fillstyle") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".FillStyle")) variable.Value = ssSTFillStyle; else variable.Optimized = true;
				variable.SetFieldName("fillstyle");
			}
			if (variable.Found && tail != null) variable.EvaluateFields(this, head, tail);
		}

		public bool ChangedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public bool OptimizedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public object AttributeGet(GlobalObjectKey key) {
			if (key == IdFillStyle) {
				return ssSTFillStyle;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssSTFillStyle.FillFromOther((IRecord) other.AttributeGet(IdFillStyle));
		}
		public bool IsDefault() {
			RCFillStyleRecord defaultStruct = new RCFillStyleRecord(null);
			if (this.ssSTFillStyle != defaultStruct.ssSTFillStyle) return false;
			return true;
		}
	} // RCFillStyleRecord

	/// <summary>
	/// Structure <code>RCFontStyleRecord</code>
	/// </summary>
	[Serializable()]
	public partial struct RCFontStyleRecord: ISerializable, ITypedRecord<RCFontStyleRecord> {
		internal static readonly GlobalObjectKey IdFontStyle = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*T1EQYMn9flehlt5R1AdK6w");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("FontStyle")]
		public STFontStyleStructure ssSTFontStyle;


		public static implicit operator STFontStyleStructure(RCFontStyleRecord r) {
			return r.ssSTFontStyle;
		}

		public static implicit operator RCFontStyleRecord(STFontStyleStructure r) {
			RCFontStyleRecord res = new RCFontStyleRecord(null);
			res.ssSTFontStyle = r;
			return res;
		}

		public BitArray OptimizedAttributes;

		public RCFontStyleRecord(params string[] dummy) {
			OptimizedAttributes = null;
			ssSTFontStyle = new STFontStyleStructure(null);
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[1];
			all[0] = null;
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
					ssSTFontStyle.OptimizedAttributes = value[0];
				}
			}
			get {
				BitArray[] all = new BitArray[1];
				all[0] = null;
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssSTFontStyle.Read(r, ref index);
		}
		/// <summary>
		/// Read from database
		/// </summary>
		/// <param name="r"> Data reader</param>
		public void ReadDB(IDataReader r) {
			int index = 0;
			Read(r, ref index);
		}

		/// <summary>
		/// Read from record
		/// </summary>
		/// <param name="r"> Record</param>
		public void ReadIM(RCFontStyleRecord r) {
			this = r;
		}


		public static bool operator == (RCFontStyleRecord a, RCFontStyleRecord b) {
			if (a.ssSTFontStyle != b.ssSTFontStyle) return false;
			return true;
		}

		public static bool operator != (RCFontStyleRecord a, RCFontStyleRecord b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(RCFontStyleRecord)) return false;
			return (this == (RCFontStyleRecord) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssSTFontStyle.GetHashCode()
				;
			} catch {
				return base.GetHashCode();
			}
		}

		public void GetObjectData(SerializationInfo info, StreamingContext context) {
			Type objInfo = this.GetType();
			FieldInfo[] fields;
			fields = objInfo.GetFields(BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			for (int i = 0; i < fields.Length; i++)
			if (fields[i] .FieldType.IsSerializable)
			info.AddValue(fields[i] .Name, fields[i] .GetValue(this));
		}

		public RCFontStyleRecord(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssSTFontStyle = new STFontStyleStructure(null);
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssSTFontStyle", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssSTFontStyle' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssSTFontStyle = (STFontStyleStructure) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssSTFontStyle.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssSTFontStyle.InternalRecursiveSave();
		}


		public RCFontStyleRecord Duplicate() {
			RCFontStyleRecord t;
			t.ssSTFontStyle = (STFontStyleStructure) this.ssSTFontStyle.Duplicate();
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Record");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
			}
			if (detailLevel > 0) {
				ssSTFontStyle.ToXml(this, recordElem, "FontStyle", detailLevel - 1);
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "fontstyle") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".FontStyle")) variable.Value = ssSTFontStyle; else variable.Optimized = true;
				variable.SetFieldName("fontstyle");
			}
			if (variable.Found && tail != null) variable.EvaluateFields(this, head, tail);
		}

		public bool ChangedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public bool OptimizedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public object AttributeGet(GlobalObjectKey key) {
			if (key == IdFontStyle) {
				return ssSTFontStyle;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssSTFontStyle.FillFromOther((IRecord) other.AttributeGet(IdFontStyle));
		}
		public bool IsDefault() {
			RCFontStyleRecord defaultStruct = new RCFontStyleRecord(null);
			if (this.ssSTFontStyle != defaultStruct.ssSTFontStyle) return false;
			return true;
		}
	} // RCFontStyleRecord

	/// <summary>
	/// Structure <code>RCNewSheetRecord</code>
	/// </summary>
	[Serializable()]
	public partial struct RCNewSheetRecord: ISerializable, ITypedRecord<RCNewSheetRecord> {
		internal static readonly GlobalObjectKey IdNewSheet = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*8PN_BTMFTs6HsuytG2X1sg");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("NewSheet")]
		public STNewSheetStructure ssSTNewSheet;


		public static implicit operator STNewSheetStructure(RCNewSheetRecord r) {
			return r.ssSTNewSheet;
		}

		public static implicit operator RCNewSheetRecord(STNewSheetStructure r) {
			RCNewSheetRecord res = new RCNewSheetRecord(null);
			res.ssSTNewSheet = r;
			return res;
		}

		public BitArray OptimizedAttributes;

		public RCNewSheetRecord(params string[] dummy) {
			OptimizedAttributes = null;
			ssSTNewSheet = new STNewSheetStructure(null);
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[1];
			all[0] = null;
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
					ssSTNewSheet.OptimizedAttributes = value[0];
				}
			}
			get {
				BitArray[] all = new BitArray[1];
				all[0] = null;
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssSTNewSheet.Read(r, ref index);
		}
		/// <summary>
		/// Read from database
		/// </summary>
		/// <param name="r"> Data reader</param>
		public void ReadDB(IDataReader r) {
			int index = 0;
			Read(r, ref index);
		}

		/// <summary>
		/// Read from record
		/// </summary>
		/// <param name="r"> Record</param>
		public void ReadIM(RCNewSheetRecord r) {
			this = r;
		}


		public static bool operator == (RCNewSheetRecord a, RCNewSheetRecord b) {
			if (a.ssSTNewSheet != b.ssSTNewSheet) return false;
			return true;
		}

		public static bool operator != (RCNewSheetRecord a, RCNewSheetRecord b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(RCNewSheetRecord)) return false;
			return (this == (RCNewSheetRecord) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssSTNewSheet.GetHashCode()
				;
			} catch {
				return base.GetHashCode();
			}
		}

		public void GetObjectData(SerializationInfo info, StreamingContext context) {
			Type objInfo = this.GetType();
			FieldInfo[] fields;
			fields = objInfo.GetFields(BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			for (int i = 0; i < fields.Length; i++)
			if (fields[i] .FieldType.IsSerializable)
			info.AddValue(fields[i] .Name, fields[i] .GetValue(this));
		}

		public RCNewSheetRecord(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssSTNewSheet = new STNewSheetStructure(null);
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssSTNewSheet", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssSTNewSheet' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssSTNewSheet = (STNewSheetStructure) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssSTNewSheet.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssSTNewSheet.InternalRecursiveSave();
		}


		public RCNewSheetRecord Duplicate() {
			RCNewSheetRecord t;
			t.ssSTNewSheet = (STNewSheetStructure) this.ssSTNewSheet.Duplicate();
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Record");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
			}
			if (detailLevel > 0) {
				ssSTNewSheet.ToXml(this, recordElem, "NewSheet", detailLevel - 1);
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "newsheet") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".NewSheet")) variable.Value = ssSTNewSheet; else variable.Optimized = true;
				variable.SetFieldName("newsheet");
			}
			if (variable.Found && tail != null) variable.EvaluateFields(this, head, tail);
		}

		public bool ChangedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public bool OptimizedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public object AttributeGet(GlobalObjectKey key) {
			if (key == IdNewSheet) {
				return ssSTNewSheet;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssSTNewSheet.FillFromOther((IRecord) other.AttributeGet(IdNewSheet));
		}
		public bool IsDefault() {
			RCNewSheetRecord defaultStruct = new RCNewSheetRecord(null);
			if (this.ssSTNewSheet != defaultStruct.ssSTNewSheet) return false;
			return true;
		}
	} // RCNewSheetRecord

	/// <summary>
	/// Structure <code>RCImageRecord</code>
	/// </summary>
	[Serializable()]
	public partial struct RCImageRecord: ISerializable, ITypedRecord<RCImageRecord> {
		internal static readonly GlobalObjectKey IdImage = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*L8pc1RE5s9y+lvB+05yOfg");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("Image")]
		public STImageStructure ssSTImage;


		public static implicit operator STImageStructure(RCImageRecord r) {
			return r.ssSTImage;
		}

		public static implicit operator RCImageRecord(STImageStructure r) {
			RCImageRecord res = new RCImageRecord(null);
			res.ssSTImage = r;
			return res;
		}

		public BitArray OptimizedAttributes;

		public RCImageRecord(params string[] dummy) {
			OptimizedAttributes = null;
			ssSTImage = new STImageStructure(null);
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[1];
			all[0] = null;
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
					ssSTImage.OptimizedAttributes = value[0];
				}
			}
			get {
				BitArray[] all = new BitArray[1];
				all[0] = null;
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssSTImage.Read(r, ref index);
		}
		/// <summary>
		/// Read from database
		/// </summary>
		/// <param name="r"> Data reader</param>
		public void ReadDB(IDataReader r) {
			int index = 0;
			Read(r, ref index);
		}

		/// <summary>
		/// Read from record
		/// </summary>
		/// <param name="r"> Record</param>
		public void ReadIM(RCImageRecord r) {
			this = r;
		}


		public static bool operator == (RCImageRecord a, RCImageRecord b) {
			if (a.ssSTImage != b.ssSTImage) return false;
			return true;
		}

		public static bool operator != (RCImageRecord a, RCImageRecord b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(RCImageRecord)) return false;
			return (this == (RCImageRecord) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssSTImage.GetHashCode()
				;
			} catch {
				return base.GetHashCode();
			}
		}

		public void GetObjectData(SerializationInfo info, StreamingContext context) {
			Type objInfo = this.GetType();
			FieldInfo[] fields;
			fields = objInfo.GetFields(BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			for (int i = 0; i < fields.Length; i++)
			if (fields[i] .FieldType.IsSerializable)
			info.AddValue(fields[i] .Name, fields[i] .GetValue(this));
		}

		public RCImageRecord(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssSTImage = new STImageStructure(null);
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssSTImage", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssSTImage' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssSTImage = (STImageStructure) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssSTImage.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssSTImage.InternalRecursiveSave();
		}


		public RCImageRecord Duplicate() {
			RCImageRecord t;
			t.ssSTImage = (STImageStructure) this.ssSTImage.Duplicate();
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Record");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
			}
			if (detailLevel > 0) {
				ssSTImage.ToXml(this, recordElem, "Image", detailLevel - 1);
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "image") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Image")) variable.Value = ssSTImage; else variable.Optimized = true;
				variable.SetFieldName("image");
			}
			if (variable.Found && tail != null) variable.EvaluateFields(this, head, tail);
		}

		public bool ChangedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public bool OptimizedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public object AttributeGet(GlobalObjectKey key) {
			if (key == IdImage) {
				return ssSTImage;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssSTImage.FillFromOther((IRecord) other.AttributeGet(IdImage));
		}
		public bool IsDefault() {
			RCImageRecord defaultStruct = new RCImageRecord(null);
			if (this.ssSTImage != defaultStruct.ssSTImage) return false;
			return true;
		}
	} // RCImageRecord

	/// <summary>
	/// Structure <code>RCProtectionRecord</code>
	/// </summary>
	[Serializable()]
	public partial struct RCProtectionRecord: ISerializable, ITypedRecord<RCProtectionRecord> {
		internal static readonly GlobalObjectKey IdProtection = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*+rf7_yBpNP4rpi8ZD0RRGg");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("Protection")]
		public STProtectionStructure ssSTProtection;


		public static implicit operator STProtectionStructure(RCProtectionRecord r) {
			return r.ssSTProtection;
		}

		public static implicit operator RCProtectionRecord(STProtectionStructure r) {
			RCProtectionRecord res = new RCProtectionRecord(null);
			res.ssSTProtection = r;
			return res;
		}

		public BitArray OptimizedAttributes;

		public RCProtectionRecord(params string[] dummy) {
			OptimizedAttributes = null;
			ssSTProtection = new STProtectionStructure(null);
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[1];
			all[0] = null;
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
					ssSTProtection.OptimizedAttributes = value[0];
				}
			}
			get {
				BitArray[] all = new BitArray[1];
				all[0] = null;
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssSTProtection.Read(r, ref index);
		}
		/// <summary>
		/// Read from database
		/// </summary>
		/// <param name="r"> Data reader</param>
		public void ReadDB(IDataReader r) {
			int index = 0;
			Read(r, ref index);
		}

		/// <summary>
		/// Read from record
		/// </summary>
		/// <param name="r"> Record</param>
		public void ReadIM(RCProtectionRecord r) {
			this = r;
		}


		public static bool operator == (RCProtectionRecord a, RCProtectionRecord b) {
			if (a.ssSTProtection != b.ssSTProtection) return false;
			return true;
		}

		public static bool operator != (RCProtectionRecord a, RCProtectionRecord b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(RCProtectionRecord)) return false;
			return (this == (RCProtectionRecord) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssSTProtection.GetHashCode()
				;
			} catch {
				return base.GetHashCode();
			}
		}

		public void GetObjectData(SerializationInfo info, StreamingContext context) {
			Type objInfo = this.GetType();
			FieldInfo[] fields;
			fields = objInfo.GetFields(BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			for (int i = 0; i < fields.Length; i++)
			if (fields[i] .FieldType.IsSerializable)
			info.AddValue(fields[i] .Name, fields[i] .GetValue(this));
		}

		public RCProtectionRecord(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssSTProtection = new STProtectionStructure(null);
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssSTProtection", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssSTProtection' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssSTProtection = (STProtectionStructure) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssSTProtection.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssSTProtection.InternalRecursiveSave();
		}


		public RCProtectionRecord Duplicate() {
			RCProtectionRecord t;
			t.ssSTProtection = (STProtectionStructure) this.ssSTProtection.Duplicate();
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Record");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
			}
			if (detailLevel > 0) {
				ssSTProtection.ToXml(this, recordElem, "Protection", detailLevel - 1);
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "protection") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Protection")) variable.Value = ssSTProtection; else variable.Optimized = true;
				variable.SetFieldName("protection");
			}
			if (variable.Found && tail != null) variable.EvaluateFields(this, head, tail);
		}

		public bool ChangedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public bool OptimizedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public object AttributeGet(GlobalObjectKey key) {
			if (key == IdProtection) {
				return ssSTProtection;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssSTProtection.FillFromOther((IRecord) other.AttributeGet(IdProtection));
		}
		public bool IsDefault() {
			RCProtectionRecord defaultStruct = new RCProtectionRecord(null);
			if (this.ssSTProtection != defaultStruct.ssSTProtection) return false;
			return true;
		}
	} // RCProtectionRecord

	/// <summary>
	/// Structure <code>RCValueRecord</code>
	/// </summary>
	[Serializable()]
	public partial struct RCValueRecord: ISerializable, ITypedRecord<RCValueRecord> {
		internal static readonly GlobalObjectKey IdValue = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*ywknZwsbBd4S4BrGWiFq7g");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("Value")]
		public STValueStructure ssSTValue;


		public static implicit operator STValueStructure(RCValueRecord r) {
			return r.ssSTValue;
		}

		public static implicit operator RCValueRecord(STValueStructure r) {
			RCValueRecord res = new RCValueRecord(null);
			res.ssSTValue = r;
			return res;
		}

		public BitArray OptimizedAttributes;

		public RCValueRecord(params string[] dummy) {
			OptimizedAttributes = null;
			ssSTValue = new STValueStructure(null);
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[1];
			all[0] = null;
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
					ssSTValue.OptimizedAttributes = value[0];
				}
			}
			get {
				BitArray[] all = new BitArray[1];
				all[0] = null;
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssSTValue.Read(r, ref index);
		}
		/// <summary>
		/// Read from database
		/// </summary>
		/// <param name="r"> Data reader</param>
		public void ReadDB(IDataReader r) {
			int index = 0;
			Read(r, ref index);
		}

		/// <summary>
		/// Read from record
		/// </summary>
		/// <param name="r"> Record</param>
		public void ReadIM(RCValueRecord r) {
			this = r;
		}


		public static bool operator == (RCValueRecord a, RCValueRecord b) {
			if (a.ssSTValue != b.ssSTValue) return false;
			return true;
		}

		public static bool operator != (RCValueRecord a, RCValueRecord b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(RCValueRecord)) return false;
			return (this == (RCValueRecord) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssSTValue.GetHashCode()
				;
			} catch {
				return base.GetHashCode();
			}
		}

		public void GetObjectData(SerializationInfo info, StreamingContext context) {
			Type objInfo = this.GetType();
			FieldInfo[] fields;
			fields = objInfo.GetFields(BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			for (int i = 0; i < fields.Length; i++)
			if (fields[i] .FieldType.IsSerializable)
			info.AddValue(fields[i] .Name, fields[i] .GetValue(this));
		}

		public RCValueRecord(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssSTValue = new STValueStructure(null);
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssSTValue", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssSTValue' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssSTValue = (STValueStructure) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssSTValue.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssSTValue.InternalRecursiveSave();
		}


		public RCValueRecord Duplicate() {
			RCValueRecord t;
			t.ssSTValue = (STValueStructure) this.ssSTValue.Duplicate();
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Record");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
			}
			if (detailLevel > 0) {
				ssSTValue.ToXml(this, recordElem, "Value", detailLevel - 1);
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "value") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Value")) variable.Value = ssSTValue; else variable.Optimized = true;
				variable.SetFieldName("value");
			}
			if (variable.Found && tail != null) variable.EvaluateFields(this, head, tail);
		}

		public bool ChangedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public bool OptimizedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public object AttributeGet(GlobalObjectKey key) {
			if (key == IdValue) {
				return ssSTValue;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssSTValue.FillFromOther((IRecord) other.AttributeGet(IdValue));
		}
		public bool IsDefault() {
			RCValueRecord defaultStruct = new RCValueRecord(null);
			if (this.ssSTValue != defaultStruct.ssSTValue) return false;
			return true;
		}
	} // RCValueRecord

	/// <summary>
	/// Structure <code>RCItemsRecord</code>
	/// </summary>
	[Serializable()]
	public partial struct RCItemsRecord: ISerializable, ITypedRecord<RCItemsRecord> {
		internal static readonly GlobalObjectKey IdItems = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*AIELN5_EkS2OyfI9cpduHw");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("Items")]
		public STItemsStructure ssSTItems;


		public static implicit operator STItemsStructure(RCItemsRecord r) {
			return r.ssSTItems;
		}

		public static implicit operator RCItemsRecord(STItemsStructure r) {
			RCItemsRecord res = new RCItemsRecord(null);
			res.ssSTItems = r;
			return res;
		}

		public BitArray OptimizedAttributes;

		public RCItemsRecord(params string[] dummy) {
			OptimizedAttributes = null;
			ssSTItems = new STItemsStructure(null);
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[1];
			all[0] = null;
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
					ssSTItems.OptimizedAttributes = value[0];
				}
			}
			get {
				BitArray[] all = new BitArray[1];
				all[0] = null;
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssSTItems.Read(r, ref index);
		}
		/// <summary>
		/// Read from database
		/// </summary>
		/// <param name="r"> Data reader</param>
		public void ReadDB(IDataReader r) {
			int index = 0;
			Read(r, ref index);
		}

		/// <summary>
		/// Read from record
		/// </summary>
		/// <param name="r"> Record</param>
		public void ReadIM(RCItemsRecord r) {
			this = r;
		}


		public static bool operator == (RCItemsRecord a, RCItemsRecord b) {
			if (a.ssSTItems != b.ssSTItems) return false;
			return true;
		}

		public static bool operator != (RCItemsRecord a, RCItemsRecord b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(RCItemsRecord)) return false;
			return (this == (RCItemsRecord) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssSTItems.GetHashCode()
				;
			} catch {
				return base.GetHashCode();
			}
		}

		public void GetObjectData(SerializationInfo info, StreamingContext context) {
			Type objInfo = this.GetType();
			FieldInfo[] fields;
			fields = objInfo.GetFields(BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			for (int i = 0; i < fields.Length; i++)
			if (fields[i] .FieldType.IsSerializable)
			info.AddValue(fields[i] .Name, fields[i] .GetValue(this));
		}

		public RCItemsRecord(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssSTItems = new STItemsStructure(null);
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssSTItems", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssSTItems' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssSTItems = (STItemsStructure) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssSTItems.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssSTItems.InternalRecursiveSave();
		}


		public RCItemsRecord Duplicate() {
			RCItemsRecord t;
			t.ssSTItems = (STItemsStructure) this.ssSTItems.Duplicate();
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Record");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
			}
			if (detailLevel > 0) {
				ssSTItems.ToXml(this, recordElem, "Items", detailLevel - 1);
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "items") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Items")) variable.Value = ssSTItems; else variable.Optimized = true;
				variable.SetFieldName("items");
			}
			if (variable.Found && tail != null) variable.EvaluateFields(this, head, tail);
		}

		public bool ChangedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public bool OptimizedAttributeGet(GlobalObjectKey key) {
			throw new Exception("Method not Supported");
		}

		public object AttributeGet(GlobalObjectKey key) {
			if (key == IdItems) {
				return ssSTItems;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssSTItems.FillFromOther((IRecord) other.AttributeGet(IdItems));
		}
		public bool IsDefault() {
			RCItemsRecord defaultStruct = new RCItemsRecord(null);
			if (this.ssSTItems != defaultStruct.ssSTItems) return false;
			return true;
		}
	} // RCItemsRecord
}
