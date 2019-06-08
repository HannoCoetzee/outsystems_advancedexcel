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
		private static readonly GlobalObjectKey IdCellFormat = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*_Kt8x60uod_oXugTPblTIQ");

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
		private static readonly GlobalObjectKey IdWorkbook = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*cSfFgXnOmYAmiyqaFCX8Wg");

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
		private static readonly GlobalObjectKey IdWorksheet = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*33h7wXL3Z32j+S7n4JJ2+g");

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
		private static readonly GlobalObjectKey IdColor = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*egnk0hJuQI_VWFqkbK8pLw");

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
		private static readonly GlobalObjectKey IdDimension = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*GgUWI0Z2l9Rs7FF+CSAoGQ");

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
		private static readonly GlobalObjectKey IdAddress = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*sakV9VS1OspLz+KlcvXSag");

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
		private static readonly GlobalObjectKey IdRange = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*fkdXsiMILofCapOVw0TaOg");

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
		private static readonly GlobalObjectKey IdDataSeries = GlobalObjectKey.Parse("2UmDmepsh0WSfJ_D1JexCA*8aE7EofndyiSFdNxWsJcng");

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
}
