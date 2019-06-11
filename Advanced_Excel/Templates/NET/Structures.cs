using System;
using System.Collections;
using System.Data;
using System.Reflection;
using System.Runtime.Serialization;
using OutSystems.ObjectKeys;
using OutSystems.RuntimeCommon;
using OutSystems.HubEdition.RuntimePlatform;
using OutSystems.HubEdition.RuntimePlatform.Db;
using OutSystems.Internal.Db;

namespace OutSystems.NssAdvanced_Excel {

	/// <summary>
	/// Structure <code>STCellFormatStructure</code> that represents the Service Studio structure
	///  <code>CellFormat</code> <p> Description: Structure to define all the formatting attributes that ca
	/// n apply to a cell.</p>
	/// </summary>
	[Serializable()]
	public partial struct STCellFormatStructure: ISerializable, ITypedRecord<STCellFormatStructure>, ISimpleRecord {
		private static readonly GlobalObjectKey IdFontName = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*BycYa7Www0ikxqPdTRgbGw");
		private static readonly GlobalObjectKey IdFontSize = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*fQrThGBwgUybjIyLimeOPA");
		private static readonly GlobalObjectKey IdBackgroundColor = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*3oXRIbtWvkKxhz4Cx7wIRg");
		private static readonly GlobalObjectKey IdAutofitColumn = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*fH82ZL_ULky72619BBkR+w");
		private static readonly GlobalObjectKey IdBold = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*GU0F2kjUe0W09TXtoXF07g");
		private static readonly GlobalObjectKey IdNumberFormat = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*T9IMXF6dHkeMjRsrROBwig");
		private static readonly GlobalObjectKey IdBorderStyle = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*4nZdbtDhvk+RtDYUBpdevA");
		private static readonly GlobalObjectKey IdBorderColor = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*ZTYcPbwyLEuVyw+rm2CbRg");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("FontName")]
		public string ssFontName;

		[System.Xml.Serialization.XmlElement("FontSize")]
		public int ssFontSize;

		[System.Xml.Serialization.XmlElement("BackgroundColor")]
		public string ssBackgroundColor;

		[System.Xml.Serialization.XmlElement("AutofitColumn")]
		public bool ssAutofitColumn;

		[System.Xml.Serialization.XmlElement("Bold")]
		public bool ssBold;

		[System.Xml.Serialization.XmlElement("NumberFormat")]
		public string ssNumberFormat;

		[System.Xml.Serialization.XmlElement("BorderStyle")]
		public int ssBorderStyle;

		[System.Xml.Serialization.XmlElement("BorderColor")]
		public string ssBorderColor;


		public BitArray OptimizedAttributes;

		public STCellFormatStructure(params string[] dummy) {
			OptimizedAttributes = null;
			ssFontName = "";
			ssFontSize = 0;
			ssBackgroundColor = "";
			ssAutofitColumn = false;
			ssBold = false;
			ssNumberFormat = "";
			ssBorderStyle = 0;
			ssBorderColor = "";
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[0];
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
				}
			}
			get {
				BitArray[] all = new BitArray[0];
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssFontName = r.ReadText(index++, "CellFormat.FontName", "");
			ssFontSize = r.ReadInteger(index++, "CellFormat.FontSize", 0);
			ssBackgroundColor = r.ReadText(index++, "CellFormat.BackgroundColor", "");
			ssAutofitColumn = r.ReadBoolean(index++, "CellFormat.AutofitColumn", false);
			ssBold = r.ReadBoolean(index++, "CellFormat.Bold", false);
			ssNumberFormat = r.ReadText(index++, "CellFormat.NumberFormat", "");
			ssBorderStyle = r.ReadInteger(index++, "CellFormat.BorderStyle", 0);
			ssBorderColor = r.ReadText(index++, "CellFormat.BorderColor", "");
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
		public void ReadIM(STCellFormatStructure r) {
			this = r;
		}


		public static bool operator == (STCellFormatStructure a, STCellFormatStructure b) {
			if (a.ssFontName != b.ssFontName) return false;
			if (a.ssFontSize != b.ssFontSize) return false;
			if (a.ssBackgroundColor != b.ssBackgroundColor) return false;
			if (a.ssAutofitColumn != b.ssAutofitColumn) return false;
			if (a.ssBold != b.ssBold) return false;
			if (a.ssNumberFormat != b.ssNumberFormat) return false;
			if (a.ssBorderStyle != b.ssBorderStyle) return false;
			if (a.ssBorderColor != b.ssBorderColor) return false;
			return true;
		}

		public static bool operator != (STCellFormatStructure a, STCellFormatStructure b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(STCellFormatStructure)) return false;
			return (this == (STCellFormatStructure) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssFontName.GetHashCode()
				^ ssFontSize.GetHashCode()
				^ ssBackgroundColor.GetHashCode()
				^ ssAutofitColumn.GetHashCode()
				^ ssBold.GetHashCode()
				^ ssNumberFormat.GetHashCode()
				^ ssBorderStyle.GetHashCode()
				^ ssBorderColor.GetHashCode()
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

		public STCellFormatStructure(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssFontName = "";
			ssFontSize = 0;
			ssBackgroundColor = "";
			ssAutofitColumn = false;
			ssBold = false;
			ssNumberFormat = "";
			ssBorderStyle = 0;
			ssBorderColor = "";
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssFontName", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssFontName' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssFontName = (string) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssFontSize", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssFontSize' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssFontSize = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssBackgroundColor", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssBackgroundColor' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssBackgroundColor = (string) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssAutofitColumn", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssAutofitColumn' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssAutofitColumn = (bool) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssBold", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssBold' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssBold = (bool) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssNumberFormat", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssNumberFormat' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssNumberFormat = (string) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssBorderStyle", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssBorderStyle' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssBorderStyle = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssBorderColor", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssBorderColor' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssBorderColor = (string) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
		}

		public void InternalRecursiveSave() {
		}


		public STCellFormatStructure Duplicate() {
			STCellFormatStructure t;
			t.ssFontName = this.ssFontName;
			t.ssFontSize = this.ssFontSize;
			t.ssBackgroundColor = this.ssBackgroundColor;
			t.ssAutofitColumn = this.ssAutofitColumn;
			t.ssBold = this.ssBold;
			t.ssNumberFormat = this.ssNumberFormat;
			t.ssBorderStyle = this.ssBorderStyle;
			t.ssBorderColor = this.ssBorderColor;
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Structure");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
				fieldName = fieldName.ToLowerInvariant();
			}
			if (detailLevel > 0) {
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".FontName")) VarValue.AppendAttribute(recordElem, "FontName", ssFontName, detailLevel, TypeKind.Text); else VarValue.AppendOptimizedAttribute(recordElem, "FontName");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".FontSize")) VarValue.AppendAttribute(recordElem, "FontSize", ssFontSize, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "FontSize");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".BackgroundColor")) VarValue.AppendAttribute(recordElem, "BackgroundColor", ssBackgroundColor, detailLevel, TypeKind.Text); else VarValue.AppendOptimizedAttribute(recordElem, "BackgroundColor");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".AutofitColumn")) VarValue.AppendAttribute(recordElem, "AutofitColumn", ssAutofitColumn, detailLevel, TypeKind.Boolean); else VarValue.AppendOptimizedAttribute(recordElem, "AutofitColumn");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Bold")) VarValue.AppendAttribute(recordElem, "Bold", ssBold, detailLevel, TypeKind.Boolean); else VarValue.AppendOptimizedAttribute(recordElem, "Bold");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".NumberFormat")) VarValue.AppendAttribute(recordElem, "NumberFormat", ssNumberFormat, detailLevel, TypeKind.Text); else VarValue.AppendOptimizedAttribute(recordElem, "NumberFormat");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".BorderStyle")) VarValue.AppendAttribute(recordElem, "BorderStyle", ssBorderStyle, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "BorderStyle");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".BorderColor")) VarValue.AppendAttribute(recordElem, "BorderColor", ssBorderColor, detailLevel, TypeKind.Text); else VarValue.AppendOptimizedAttribute(recordElem, "BorderColor");
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "fontname") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".FontName")) variable.Value = ssFontName; else variable.Optimized = true;
			} else if (head == "fontsize") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".FontSize")) variable.Value = ssFontSize; else variable.Optimized = true;
			} else if (head == "backgroundcolor") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".BackgroundColor")) variable.Value = ssBackgroundColor; else variable.Optimized = true;
			} else if (head == "autofitcolumn") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".AutofitColumn")) variable.Value = ssAutofitColumn; else variable.Optimized = true;
			} else if (head == "bold") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Bold")) variable.Value = ssBold; else variable.Optimized = true;
			} else if (head == "numberformat") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".NumberFormat")) variable.Value = ssNumberFormat; else variable.Optimized = true;
			} else if (head == "borderstyle") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".BorderStyle")) variable.Value = ssBorderStyle; else variable.Optimized = true;
			} else if (head == "bordercolor") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".BorderColor")) variable.Value = ssBorderColor; else variable.Optimized = true;
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
			if (key == IdFontName) {
				return ssFontName;
			} else if (key == IdFontSize) {
				return ssFontSize;
			} else if (key == IdBackgroundColor) {
				return ssBackgroundColor;
			} else if (key == IdAutofitColumn) {
				return ssAutofitColumn;
			} else if (key == IdBold) {
				return ssBold;
			} else if (key == IdNumberFormat) {
				return ssNumberFormat;
			} else if (key == IdBorderStyle) {
				return ssBorderStyle;
			} else if (key == IdBorderColor) {
				return ssBorderColor;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssFontName = (string) other.AttributeGet(IdFontName);
			ssFontSize = (int) other.AttributeGet(IdFontSize);
			ssBackgroundColor = (string) other.AttributeGet(IdBackgroundColor);
			ssAutofitColumn = (bool) other.AttributeGet(IdAutofitColumn);
			ssBold = (bool) other.AttributeGet(IdBold);
			ssNumberFormat = (string) other.AttributeGet(IdNumberFormat);
			ssBorderStyle = (int) other.AttributeGet(IdBorderStyle);
			ssBorderColor = (string) other.AttributeGet(IdBorderColor);
		}
		public bool IsDefault() {
			STCellFormatStructure defaultStruct = new STCellFormatStructure(null);
			if (this.ssFontName != defaultStruct.ssFontName) return false;
			if (this.ssFontSize != defaultStruct.ssFontSize) return false;
			if (this.ssBackgroundColor != defaultStruct.ssBackgroundColor) return false;
			if (this.ssAutofitColumn != defaultStruct.ssAutofitColumn) return false;
			if (this.ssBold != defaultStruct.ssBold) return false;
			if (this.ssNumberFormat != defaultStruct.ssNumberFormat) return false;
			if (this.ssBorderStyle != defaultStruct.ssBorderStyle) return false;
			if (this.ssBorderColor != defaultStruct.ssBorderColor) return false;
			return true;
		}
	} // STCellFormatStructure

	/// <summary>
	/// Structure <code>STWorkbookStructure</code> that represents the Service Studio structure
	///  <code>Workbook</code> <p> Description: The Excel File</p>
	/// </summary>
	[Serializable()]
	public partial struct STWorkbookStructure: ISerializable, ITypedRecord<STWorkbookStructure>, ISimpleRecord {
		private static readonly GlobalObjectKey IdWorksheets = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*izrPaRKZ_EqpgPLXabozRQ");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("Worksheets")]
		public RLWorksheetRecordList ssWorksheets;


		public BitArray OptimizedAttributes;

		public STWorkbookStructure(params string[] dummy) {
			OptimizedAttributes = null;
			ssWorksheets = new RLWorksheetRecordList();
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[0];
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
				}
			}
			get {
				BitArray[] all = new BitArray[0];
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
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
		public void ReadIM(STWorkbookStructure r) {
			this = r;
		}


		public static bool operator == (STWorkbookStructure a, STWorkbookStructure b) {
			if (a.ssWorksheets != b.ssWorksheets) return false;
			return true;
		}

		public static bool operator != (STWorkbookStructure a, STWorkbookStructure b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(STWorkbookStructure)) return false;
			return (this == (STWorkbookStructure) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssWorksheets.GetHashCode()
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

		public STWorkbookStructure(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssWorksheets = new RLWorksheetRecordList();
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssWorksheets", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssWorksheets' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssWorksheets = (RLWorksheetRecordList) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssWorksheets.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssWorksheets.InternalRecursiveSave();
		}


		public STWorkbookStructure Duplicate() {
			STWorkbookStructure t;
			t.ssWorksheets = (RLWorksheetRecordList) this.ssWorksheets.Duplicate();
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Structure");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
				fieldName = fieldName.ToLowerInvariant();
			}
			if (detailLevel > 0) {
				ssWorksheets.ToXml(this, recordElem, "Worksheets", detailLevel - 1);
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "worksheets") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Worksheets")) variable.Value = ssWorksheets; else variable.Optimized = true;
				variable.SetFieldName("worksheets");
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
			if (key == IdWorksheets) {
				return ssWorksheets;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssWorksheets = new RLWorksheetRecordList();
			ssWorksheets.FillFromOther((IOSList) other.AttributeGet(IdWorksheets));
		}
		public bool IsDefault() {
			STWorkbookStructure defaultStruct = new STWorkbookStructure(null);
			if (this.ssWorksheets != null && this.ssWorksheets.Length != 0) return false;
			return true;
		}
	} // STWorkbookStructure

	/// <summary>
	/// Structure <code>STWorksheetStructure</code> that represents the Service Studio structure
	///  <code>Worksheet</code> <p> Description: Structure defining attributes pertaining to
	/// a worksheet</p>
	/// </summary>
	[Serializable()]
	public partial struct STWorksheetStructure: ISerializable, ITypedRecord<STWorksheetStructure>, ISimpleRecord {
		private static readonly GlobalObjectKey IdName = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*qu+kXpI5mk2nEoopujaw1w");
		private static readonly GlobalObjectKey IdIndex = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*y5APBAwZfkWoTqWCP01Pxg");
		private static readonly GlobalObjectKey IdTabColor = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*vXO_uSpdREW_4CDsVlbFlQ");
		private static readonly GlobalObjectKey IdDimension = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*xXgLP3ev_0mowvNYBfWMGg");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("Name")]
		public string ssName;

		[System.Xml.Serialization.XmlElement("Index")]
		public int ssIndex;

		[System.Xml.Serialization.XmlElement("TabColor")]
		public RCColorRecord ssTabColor;

		[System.Xml.Serialization.XmlElement("Dimension")]
		public RCDimensionRecord ssDimension;


		public BitArray OptimizedAttributes;

		public STWorksheetStructure(params string[] dummy) {
			OptimizedAttributes = null;
			ssName = "";
			ssIndex = 0;
			ssTabColor = new RCColorRecord(null);
			ssDimension = new RCDimensionRecord(null);
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[2];
			all[0] = null;
			all[1] = null;
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
					ssTabColor.OptimizedAttributes = value[0];
					ssDimension.OptimizedAttributes = value[1];
				}
			}
			get {
				BitArray[] all = new BitArray[2];
				all[0] = null;
				all[1] = null;
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssName = r.ReadText(index++, "Worksheet.Name", "");
			ssIndex = r.ReadInteger(index++, "Worksheet.Index", 0);
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
		public void ReadIM(STWorksheetStructure r) {
			this = r;
		}


		public static bool operator == (STWorksheetStructure a, STWorksheetStructure b) {
			if (a.ssName != b.ssName) return false;
			if (a.ssIndex != b.ssIndex) return false;
			if (a.ssTabColor != b.ssTabColor) return false;
			if (a.ssDimension != b.ssDimension) return false;
			return true;
		}

		public static bool operator != (STWorksheetStructure a, STWorksheetStructure b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(STWorksheetStructure)) return false;
			return (this == (STWorksheetStructure) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssName.GetHashCode()
				^ ssIndex.GetHashCode()
				^ ssTabColor.GetHashCode()
				^ ssDimension.GetHashCode()
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

		public STWorksheetStructure(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssName = "";
			ssIndex = 0;
			ssTabColor = new RCColorRecord(null);
			ssDimension = new RCDimensionRecord(null);
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssName", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssName' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssName = (string) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssIndex", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssIndex' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssIndex = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssTabColor", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssTabColor' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssTabColor = (RCColorRecord) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssDimension", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssDimension' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssDimension = (RCDimensionRecord) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssTabColor.RecursiveReset();
			ssDimension.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssTabColor.InternalRecursiveSave();
			ssDimension.InternalRecursiveSave();
		}


		public STWorksheetStructure Duplicate() {
			STWorksheetStructure t;
			t.ssName = this.ssName;
			t.ssIndex = this.ssIndex;
			t.ssTabColor = (RCColorRecord) this.ssTabColor.Duplicate();
			t.ssDimension = (RCDimensionRecord) this.ssDimension.Duplicate();
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Structure");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
				fieldName = fieldName.ToLowerInvariant();
			}
			if (detailLevel > 0) {
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Name")) VarValue.AppendAttribute(recordElem, "Name", ssName, detailLevel, TypeKind.Text); else VarValue.AppendOptimizedAttribute(recordElem, "Name");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Index")) VarValue.AppendAttribute(recordElem, "Index", ssIndex, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "Index");
				ssTabColor.ToXml(this, recordElem, "TabColor", detailLevel - 1);
				ssDimension.ToXml(this, recordElem, "Dimension", detailLevel - 1);
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "name") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Name")) variable.Value = ssName; else variable.Optimized = true;
			} else if (head == "index") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Index")) variable.Value = ssIndex; else variable.Optimized = true;
			} else if (head == "tabcolor") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".TabColor")) variable.Value = ssTabColor; else variable.Optimized = true;
				variable.SetFieldName("tabcolor");
			} else if (head == "dimension") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Dimension")) variable.Value = ssDimension; else variable.Optimized = true;
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
			if (key == IdName) {
				return ssName;
			} else if (key == IdIndex) {
				return ssIndex;
			} else if (key == IdTabColor) {
				return ssTabColor;
			} else if (key == IdDimension) {
				return ssDimension;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssName = (string) other.AttributeGet(IdName);
			ssIndex = (int) other.AttributeGet(IdIndex);
			ssTabColor.FillFromOther((IRecord) other.AttributeGet(IdTabColor));
			ssDimension.FillFromOther((IRecord) other.AttributeGet(IdDimension));
		}
		public bool IsDefault() {
			STWorksheetStructure defaultStruct = new STWorksheetStructure(null);
			if (this.ssName != defaultStruct.ssName) return false;
			if (this.ssIndex != defaultStruct.ssIndex) return false;
			if (this.ssTabColor != defaultStruct.ssTabColor) return false;
			if (this.ssDimension != defaultStruct.ssDimension) return false;
			return true;
		}
	} // STWorksheetStructure

	/// <summary>
	/// Structure <code>STColorStructure</code> that represents the Service Studio structure
	///  <code>Color</code> <p> Description: Represents an ARGB (alpha, red, green, blue) color.</p>
	/// </summary>
	[Serializable()]
	public partial struct STColorStructure: ISerializable, ITypedRecord<STColorStructure>, ISimpleRecord {
		private static readonly GlobalObjectKey IdIsKnownColor = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*l5cu6CGHqkSOaN_PCVYsDA");
		private static readonly GlobalObjectKey IdIsNamedColor = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*eekKwslJf0u5aYbXGPnflA");
		private static readonly GlobalObjectKey IdIsSystemColor = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*OD+P0K4xt02YijR2RL18uw");
		private static readonly GlobalObjectKey IdA = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*zUlWKdvonEGw1Rc65EoX6w");
		private static readonly GlobalObjectKey IdR = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*ECHSWe_rkkuNSf04BI6Adg");
		private static readonly GlobalObjectKey IdG = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*5YqhHvTzz0aEouuGyeT4Ww");
		private static readonly GlobalObjectKey IdB = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*c2IM4kriBEiuF+KhK3PbSg");
		private static readonly GlobalObjectKey IdName = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*eyCFq+evuk6c3ZgZ2JJS_g");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("IsKnownColor")]
		public bool ssIsKnownColor;

		[System.Xml.Serialization.XmlElement("IsNamedColor")]
		public bool ssIsNamedColor;

		[System.Xml.Serialization.XmlElement("IsSystemColor")]
		public bool ssIsSystemColor;

		[System.Xml.Serialization.XmlElement("A")]
		public int ssA;

		[System.Xml.Serialization.XmlElement("R")]
		public int ssR;

		[System.Xml.Serialization.XmlElement("G")]
		public int ssG;

		[System.Xml.Serialization.XmlElement("B")]
		public int ssB;

		[System.Xml.Serialization.XmlElement("Name")]
		public string ssName;


		public BitArray OptimizedAttributes;

		public STColorStructure(params string[] dummy) {
			OptimizedAttributes = null;
			ssIsKnownColor = false;
			ssIsNamedColor = false;
			ssIsSystemColor = false;
			ssA = 0;
			ssR = 0;
			ssG = 0;
			ssB = 0;
			ssName = "";
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[0];
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
				}
			}
			get {
				BitArray[] all = new BitArray[0];
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssIsKnownColor = r.ReadBoolean(index++, "Color.IsKnownColor", false);
			ssIsNamedColor = r.ReadBoolean(index++, "Color.IsNamedColor", false);
			ssIsSystemColor = r.ReadBoolean(index++, "Color.IsSystemColor", false);
			ssA = r.ReadInteger(index++, "Color.A", 0);
			ssR = r.ReadInteger(index++, "Color.R", 0);
			ssG = r.ReadInteger(index++, "Color.G", 0);
			ssB = r.ReadInteger(index++, "Color.B", 0);
			ssName = r.ReadText(index++, "Color.Name", "");
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
		public void ReadIM(STColorStructure r) {
			this = r;
		}


		public static bool operator == (STColorStructure a, STColorStructure b) {
			if (a.ssIsKnownColor != b.ssIsKnownColor) return false;
			if (a.ssIsNamedColor != b.ssIsNamedColor) return false;
			if (a.ssIsSystemColor != b.ssIsSystemColor) return false;
			if (a.ssA != b.ssA) return false;
			if (a.ssR != b.ssR) return false;
			if (a.ssG != b.ssG) return false;
			if (a.ssB != b.ssB) return false;
			if (a.ssName != b.ssName) return false;
			return true;
		}

		public static bool operator != (STColorStructure a, STColorStructure b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(STColorStructure)) return false;
			return (this == (STColorStructure) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssIsKnownColor.GetHashCode()
				^ ssIsNamedColor.GetHashCode()
				^ ssIsSystemColor.GetHashCode()
				^ ssA.GetHashCode()
				^ ssR.GetHashCode()
				^ ssG.GetHashCode()
				^ ssB.GetHashCode()
				^ ssName.GetHashCode()
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

		public STColorStructure(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssIsKnownColor = false;
			ssIsNamedColor = false;
			ssIsSystemColor = false;
			ssA = 0;
			ssR = 0;
			ssG = 0;
			ssB = 0;
			ssName = "";
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssIsKnownColor", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssIsKnownColor' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssIsKnownColor = (bool) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssIsNamedColor", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssIsNamedColor' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssIsNamedColor = (bool) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssIsSystemColor", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssIsSystemColor' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssIsSystemColor = (bool) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssA", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssA' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssA = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssR", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssR' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssR = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssG", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssG' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssG = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssB", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssB' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssB = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssName", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssName' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssName = (string) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
		}

		public void InternalRecursiveSave() {
		}


		public STColorStructure Duplicate() {
			STColorStructure t;
			t.ssIsKnownColor = this.ssIsKnownColor;
			t.ssIsNamedColor = this.ssIsNamedColor;
			t.ssIsSystemColor = this.ssIsSystemColor;
			t.ssA = this.ssA;
			t.ssR = this.ssR;
			t.ssG = this.ssG;
			t.ssB = this.ssB;
			t.ssName = this.ssName;
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Structure");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
				fieldName = fieldName.ToLowerInvariant();
			}
			if (detailLevel > 0) {
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".IsKnownColor")) VarValue.AppendAttribute(recordElem, "IsKnownColor", ssIsKnownColor, detailLevel, TypeKind.Boolean); else VarValue.AppendOptimizedAttribute(recordElem, "IsKnownColor");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".IsNamedColor")) VarValue.AppendAttribute(recordElem, "IsNamedColor", ssIsNamedColor, detailLevel, TypeKind.Boolean); else VarValue.AppendOptimizedAttribute(recordElem, "IsNamedColor");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".IsSystemColor")) VarValue.AppendAttribute(recordElem, "IsSystemColor", ssIsSystemColor, detailLevel, TypeKind.Boolean); else VarValue.AppendOptimizedAttribute(recordElem, "IsSystemColor");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".A")) VarValue.AppendAttribute(recordElem, "A", ssA, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "A");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".R")) VarValue.AppendAttribute(recordElem, "R", ssR, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "R");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".G")) VarValue.AppendAttribute(recordElem, "G", ssG, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "G");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".B")) VarValue.AppendAttribute(recordElem, "B", ssB, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "B");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Name")) VarValue.AppendAttribute(recordElem, "Name", ssName, detailLevel, TypeKind.Text); else VarValue.AppendOptimizedAttribute(recordElem, "Name");
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "isknowncolor") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".IsKnownColor")) variable.Value = ssIsKnownColor; else variable.Optimized = true;
			} else if (head == "isnamedcolor") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".IsNamedColor")) variable.Value = ssIsNamedColor; else variable.Optimized = true;
			} else if (head == "issystemcolor") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".IsSystemColor")) variable.Value = ssIsSystemColor; else variable.Optimized = true;
			} else if (head == "a") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".A")) variable.Value = ssA; else variable.Optimized = true;
			} else if (head == "r") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".R")) variable.Value = ssR; else variable.Optimized = true;
			} else if (head == "g") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".G")) variable.Value = ssG; else variable.Optimized = true;
			} else if (head == "b") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".B")) variable.Value = ssB; else variable.Optimized = true;
			} else if (head == "name") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Name")) variable.Value = ssName; else variable.Optimized = true;
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
			if (key == IdIsKnownColor) {
				return ssIsKnownColor;
			} else if (key == IdIsNamedColor) {
				return ssIsNamedColor;
			} else if (key == IdIsSystemColor) {
				return ssIsSystemColor;
			} else if (key == IdA) {
				return ssA;
			} else if (key == IdR) {
				return ssR;
			} else if (key == IdG) {
				return ssG;
			} else if (key == IdB) {
				return ssB;
			} else if (key == IdName) {
				return ssName;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssIsKnownColor = (bool) other.AttributeGet(IdIsKnownColor);
			ssIsNamedColor = (bool) other.AttributeGet(IdIsNamedColor);
			ssIsSystemColor = (bool) other.AttributeGet(IdIsSystemColor);
			ssA = (int) other.AttributeGet(IdA);
			ssR = (int) other.AttributeGet(IdR);
			ssG = (int) other.AttributeGet(IdG);
			ssB = (int) other.AttributeGet(IdB);
			ssName = (string) other.AttributeGet(IdName);
		}
		public bool IsDefault() {
			STColorStructure defaultStruct = new STColorStructure(null);
			if (this.ssIsKnownColor != defaultStruct.ssIsKnownColor) return false;
			if (this.ssIsNamedColor != defaultStruct.ssIsNamedColor) return false;
			if (this.ssIsSystemColor != defaultStruct.ssIsSystemColor) return false;
			if (this.ssA != defaultStruct.ssA) return false;
			if (this.ssR != defaultStruct.ssR) return false;
			if (this.ssG != defaultStruct.ssG) return false;
			if (this.ssB != defaultStruct.ssB) return false;
			if (this.ssName != defaultStruct.ssName) return false;
			return true;
		}
	} // STColorStructure

	/// <summary>
	/// Structure <code>STDimensionStructure</code> that represents the Service Studio structure
	///  <code>Dimension</code> <p> Description: Worksheet dimension structure</p>
	/// </summary>
	[Serializable()]
	public partial struct STDimensionStructure: ISerializable, ITypedRecord<STDimensionStructure>, ISimpleRecord {
		private static readonly GlobalObjectKey IdAddress = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*yHVFLdcUqE+X718XHPI+kw");
		private static readonly GlobalObjectKey IdColumns = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*aNFKux7b1kWMQpn59TfPGQ");
		private static readonly GlobalObjectKey IdRows = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*01Ula9rHSkKShXRia3Q7sQ");
		private static readonly GlobalObjectKey IdStart = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*JaS35nzoPk2ZuTVFbuOf0Q");
		private static readonly GlobalObjectKey IdEnd = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*Gh6izWtZ+UKMWJvgHvCGSA");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("Address")]
		public string ssAddress;

		[System.Xml.Serialization.XmlElement("Columns")]
		public int ssColumns;

		[System.Xml.Serialization.XmlElement("Rows")]
		public int ssRows;

		[System.Xml.Serialization.XmlElement("Start")]
		public RCAddressRecord ssStart;

		[System.Xml.Serialization.XmlElement("End")]
		public RCAddressRecord ssEnd;


		public BitArray OptimizedAttributes;

		public STDimensionStructure(params string[] dummy) {
			OptimizedAttributes = null;
			ssAddress = "";
			ssColumns = 0;
			ssRows = 0;
			ssStart = new RCAddressRecord(null);
			ssEnd = new RCAddressRecord(null);
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[2];
			all[0] = null;
			all[1] = null;
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
					ssStart.OptimizedAttributes = value[0];
					ssEnd.OptimizedAttributes = value[1];
				}
			}
			get {
				BitArray[] all = new BitArray[2];
				all[0] = null;
				all[1] = null;
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssAddress = r.ReadText(index++, "Dimension.Address", "");
			ssColumns = r.ReadInteger(index++, "Dimension.Columns", 0);
			ssRows = r.ReadInteger(index++, "Dimension.Rows", 0);
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
		public void ReadIM(STDimensionStructure r) {
			this = r;
		}


		public static bool operator == (STDimensionStructure a, STDimensionStructure b) {
			if (a.ssAddress != b.ssAddress) return false;
			if (a.ssColumns != b.ssColumns) return false;
			if (a.ssRows != b.ssRows) return false;
			if (a.ssStart != b.ssStart) return false;
			if (a.ssEnd != b.ssEnd) return false;
			return true;
		}

		public static bool operator != (STDimensionStructure a, STDimensionStructure b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(STDimensionStructure)) return false;
			return (this == (STDimensionStructure) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssAddress.GetHashCode()
				^ ssColumns.GetHashCode()
				^ ssRows.GetHashCode()
				^ ssStart.GetHashCode()
				^ ssEnd.GetHashCode()
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

		public STDimensionStructure(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssAddress = "";
			ssColumns = 0;
			ssRows = 0;
			ssStart = new RCAddressRecord(null);
			ssEnd = new RCAddressRecord(null);
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssAddress", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssAddress' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssAddress = (string) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssColumns", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssColumns' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssColumns = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssRows", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssRows' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssRows = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssStart", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssStart' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssStart = (RCAddressRecord) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssEnd", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssEnd' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssEnd = (RCAddressRecord) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssStart.RecursiveReset();
			ssEnd.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssStart.InternalRecursiveSave();
			ssEnd.InternalRecursiveSave();
		}


		public STDimensionStructure Duplicate() {
			STDimensionStructure t;
			t.ssAddress = this.ssAddress;
			t.ssColumns = this.ssColumns;
			t.ssRows = this.ssRows;
			t.ssStart = (RCAddressRecord) this.ssStart.Duplicate();
			t.ssEnd = (RCAddressRecord) this.ssEnd.Duplicate();
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Structure");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
				fieldName = fieldName.ToLowerInvariant();
			}
			if (detailLevel > 0) {
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Address")) VarValue.AppendAttribute(recordElem, "Address", ssAddress, detailLevel, TypeKind.Text); else VarValue.AppendOptimizedAttribute(recordElem, "Address");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Columns")) VarValue.AppendAttribute(recordElem, "Columns", ssColumns, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "Columns");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Rows")) VarValue.AppendAttribute(recordElem, "Rows", ssRows, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "Rows");
				ssStart.ToXml(this, recordElem, "Start", detailLevel - 1);
				ssEnd.ToXml(this, recordElem, "End", detailLevel - 1);
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "address") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Address")) variable.Value = ssAddress; else variable.Optimized = true;
			} else if (head == "columns") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Columns")) variable.Value = ssColumns; else variable.Optimized = true;
			} else if (head == "rows") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Rows")) variable.Value = ssRows; else variable.Optimized = true;
			} else if (head == "start") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Start")) variable.Value = ssStart; else variable.Optimized = true;
				variable.SetFieldName("start");
			} else if (head == "end") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".End")) variable.Value = ssEnd; else variable.Optimized = true;
				variable.SetFieldName("end");
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
				return ssAddress;
			} else if (key == IdColumns) {
				return ssColumns;
			} else if (key == IdRows) {
				return ssRows;
			} else if (key == IdStart) {
				return ssStart;
			} else if (key == IdEnd) {
				return ssEnd;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssAddress = (string) other.AttributeGet(IdAddress);
			ssColumns = (int) other.AttributeGet(IdColumns);
			ssRows = (int) other.AttributeGet(IdRows);
			ssStart.FillFromOther((IRecord) other.AttributeGet(IdStart));
			ssEnd.FillFromOther((IRecord) other.AttributeGet(IdEnd));
		}
		public bool IsDefault() {
			STDimensionStructure defaultStruct = new STDimensionStructure(null);
			if (this.ssAddress != defaultStruct.ssAddress) return false;
			if (this.ssColumns != defaultStruct.ssColumns) return false;
			if (this.ssRows != defaultStruct.ssRows) return false;
			if (this.ssStart != defaultStruct.ssStart) return false;
			if (this.ssEnd != defaultStruct.ssEnd) return false;
			return true;
		}
	} // STDimensionStructure

	/// <summary>
	/// Structure <code>STAddressStructure</code> that represents the Service Studio structure
	///  <code>Address</code> <p> Description: Excel Address Structure</p>
	/// </summary>
	[Serializable()]
	public partial struct STAddressStructure: ISerializable, ITypedRecord<STAddressStructure>, ISimpleRecord {
		private static readonly GlobalObjectKey IdAddress = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*AVbMSVppE0+cJ1ha0PX_Tg");
		private static readonly GlobalObjectKey IdColumn = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*2K8ZTo2MUkGj+bwxIpHHSg");
		private static readonly GlobalObjectKey IdIsRef = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*q_rA6foNt02HS9grG81W9w");
		private static readonly GlobalObjectKey IdRow = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*+nGPqzJReEerUJI3NyiEPg");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("Address")]
		public string ssAddress;

		[System.Xml.Serialization.XmlElement("Column")]
		public int ssColumn;

		[System.Xml.Serialization.XmlElement("IsRef")]
		public bool ssIsRef;

		[System.Xml.Serialization.XmlElement("Row")]
		public int ssRow;


		public BitArray OptimizedAttributes;

		public STAddressStructure(params string[] dummy) {
			OptimizedAttributes = null;
			ssAddress = "";
			ssColumn = 0;
			ssIsRef = false;
			ssRow = 0;
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[0];
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
				}
			}
			get {
				BitArray[] all = new BitArray[0];
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssAddress = r.ReadText(index++, "Address.Address", "");
			ssColumn = r.ReadInteger(index++, "Address.Column", 0);
			ssIsRef = r.ReadBoolean(index++, "Address.IsRef", false);
			ssRow = r.ReadInteger(index++, "Address.Row", 0);
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
		public void ReadIM(STAddressStructure r) {
			this = r;
		}


		public static bool operator == (STAddressStructure a, STAddressStructure b) {
			if (a.ssAddress != b.ssAddress) return false;
			if (a.ssColumn != b.ssColumn) return false;
			if (a.ssIsRef != b.ssIsRef) return false;
			if (a.ssRow != b.ssRow) return false;
			return true;
		}

		public static bool operator != (STAddressStructure a, STAddressStructure b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(STAddressStructure)) return false;
			return (this == (STAddressStructure) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssAddress.GetHashCode()
				^ ssColumn.GetHashCode()
				^ ssIsRef.GetHashCode()
				^ ssRow.GetHashCode()
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

		public STAddressStructure(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssAddress = "";
			ssColumn = 0;
			ssIsRef = false;
			ssRow = 0;
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssAddress", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssAddress' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssAddress = (string) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssColumn", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssColumn' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssColumn = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssIsRef", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssIsRef' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssIsRef = (bool) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssRow", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssRow' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssRow = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
		}

		public void InternalRecursiveSave() {
		}


		public STAddressStructure Duplicate() {
			STAddressStructure t;
			t.ssAddress = this.ssAddress;
			t.ssColumn = this.ssColumn;
			t.ssIsRef = this.ssIsRef;
			t.ssRow = this.ssRow;
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Structure");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
				fieldName = fieldName.ToLowerInvariant();
			}
			if (detailLevel > 0) {
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Address")) VarValue.AppendAttribute(recordElem, "Address", ssAddress, detailLevel, TypeKind.Text); else VarValue.AppendOptimizedAttribute(recordElem, "Address");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Column")) VarValue.AppendAttribute(recordElem, "Column", ssColumn, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "Column");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".IsRef")) VarValue.AppendAttribute(recordElem, "IsRef", ssIsRef, detailLevel, TypeKind.Boolean); else VarValue.AppendOptimizedAttribute(recordElem, "IsRef");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Row")) VarValue.AppendAttribute(recordElem, "Row", ssRow, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "Row");
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "address") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Address")) variable.Value = ssAddress; else variable.Optimized = true;
			} else if (head == "column") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Column")) variable.Value = ssColumn; else variable.Optimized = true;
			} else if (head == "isref") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".IsRef")) variable.Value = ssIsRef; else variable.Optimized = true;
			} else if (head == "row") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Row")) variable.Value = ssRow; else variable.Optimized = true;
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
				return ssAddress;
			} else if (key == IdColumn) {
				return ssColumn;
			} else if (key == IdIsRef) {
				return ssIsRef;
			} else if (key == IdRow) {
				return ssRow;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssAddress = (string) other.AttributeGet(IdAddress);
			ssColumn = (int) other.AttributeGet(IdColumn);
			ssIsRef = (bool) other.AttributeGet(IdIsRef);
			ssRow = (int) other.AttributeGet(IdRow);
		}
		public bool IsDefault() {
			STAddressStructure defaultStruct = new STAddressStructure(null);
			if (this.ssAddress != defaultStruct.ssAddress) return false;
			if (this.ssColumn != defaultStruct.ssColumn) return false;
			if (this.ssIsRef != defaultStruct.ssIsRef) return false;
			if (this.ssRow != defaultStruct.ssRow) return false;
			return true;
		}
	} // STAddressStructure

	/// <summary>
	/// Structure <code>STRangeStructure</code> that represents the Service Studio structure
	///  <code>Range</code> <p> Description: Describes a range of cells from RowStart, ColStart to RowEnd
	/// , ColEnd, where 1,1 is the the 1A cell</p>
	/// </summary>
	[Serializable()]
	public partial struct STRangeStructure: ISerializable, ITypedRecord<STRangeStructure>, ISimpleRecord {
		private static readonly GlobalObjectKey IdStartRow = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*kAiDbHLyfka+BBBpQVgzag");
		private static readonly GlobalObjectKey IdStartCol = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*LjVCwwGKH02mZjrBadmzXg");
		private static readonly GlobalObjectKey IdEndRow = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*oW8H1BHvyEqUQO2kj1eHeQ");
		private static readonly GlobalObjectKey IdEndCol = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*FidsGMKZGUKVmU2u5C7LgA");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("StartRow")]
		public int ssStartRow;

		[System.Xml.Serialization.XmlElement("StartCol")]
		public int ssStartCol;

		[System.Xml.Serialization.XmlElement("EndRow")]
		public int ssEndRow;

		[System.Xml.Serialization.XmlElement("EndCol")]
		public int ssEndCol;


		public BitArray OptimizedAttributes;

		public STRangeStructure(params string[] dummy) {
			OptimizedAttributes = null;
			ssStartRow = 0;
			ssStartCol = 0;
			ssEndRow = 0;
			ssEndCol = 0;
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[0];
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
				}
			}
			get {
				BitArray[] all = new BitArray[0];
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssStartRow = r.ReadInteger(index++, "Range.StartRow", 0);
			ssStartCol = r.ReadInteger(index++, "Range.StartCol", 0);
			ssEndRow = r.ReadInteger(index++, "Range.EndRow", 0);
			ssEndCol = r.ReadInteger(index++, "Range.EndCol", 0);
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
		public void ReadIM(STRangeStructure r) {
			this = r;
		}


		public static bool operator == (STRangeStructure a, STRangeStructure b) {
			if (a.ssStartRow != b.ssStartRow) return false;
			if (a.ssStartCol != b.ssStartCol) return false;
			if (a.ssEndRow != b.ssEndRow) return false;
			if (a.ssEndCol != b.ssEndCol) return false;
			return true;
		}

		public static bool operator != (STRangeStructure a, STRangeStructure b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(STRangeStructure)) return false;
			return (this == (STRangeStructure) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssStartRow.GetHashCode()
				^ ssStartCol.GetHashCode()
				^ ssEndRow.GetHashCode()
				^ ssEndCol.GetHashCode()
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

		public STRangeStructure(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssStartRow = 0;
			ssStartCol = 0;
			ssEndRow = 0;
			ssEndCol = 0;
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssStartRow", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssStartRow' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssStartRow = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssStartCol", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssStartCol' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssStartCol = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssEndRow", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssEndRow' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssEndRow = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssEndCol", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssEndCol' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssEndCol = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
		}

		public void InternalRecursiveSave() {
		}


		public STRangeStructure Duplicate() {
			STRangeStructure t;
			t.ssStartRow = this.ssStartRow;
			t.ssStartCol = this.ssStartCol;
			t.ssEndRow = this.ssEndRow;
			t.ssEndCol = this.ssEndCol;
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Structure");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
				fieldName = fieldName.ToLowerInvariant();
			}
			if (detailLevel > 0) {
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".StartRow")) VarValue.AppendAttribute(recordElem, "StartRow", ssStartRow, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "StartRow");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".StartCol")) VarValue.AppendAttribute(recordElem, "StartCol", ssStartCol, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "StartCol");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".EndRow")) VarValue.AppendAttribute(recordElem, "EndRow", ssEndRow, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "EndRow");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".EndCol")) VarValue.AppendAttribute(recordElem, "EndCol", ssEndCol, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "EndCol");
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "startrow") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".StartRow")) variable.Value = ssStartRow; else variable.Optimized = true;
			} else if (head == "startcol") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".StartCol")) variable.Value = ssStartCol; else variable.Optimized = true;
			} else if (head == "endrow") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".EndRow")) variable.Value = ssEndRow; else variable.Optimized = true;
			} else if (head == "endcol") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".EndCol")) variable.Value = ssEndCol; else variable.Optimized = true;
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
			if (key == IdStartRow) {
				return ssStartRow;
			} else if (key == IdStartCol) {
				return ssStartCol;
			} else if (key == IdEndRow) {
				return ssEndRow;
			} else if (key == IdEndCol) {
				return ssEndCol;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssStartRow = (int) other.AttributeGet(IdStartRow);
			ssStartCol = (int) other.AttributeGet(IdStartCol);
			ssEndRow = (int) other.AttributeGet(IdEndRow);
			ssEndCol = (int) other.AttributeGet(IdEndCol);
		}
		public bool IsDefault() {
			STRangeStructure defaultStruct = new STRangeStructure(null);
			if (this.ssStartRow != defaultStruct.ssStartRow) return false;
			if (this.ssStartCol != defaultStruct.ssStartCol) return false;
			if (this.ssEndRow != defaultStruct.ssEndRow) return false;
			if (this.ssEndCol != defaultStruct.ssEndCol) return false;
			return true;
		}
	} // STRangeStructure

	/// <summary>
	/// Structure <code>STDataSeriesStructure</code> that represents the Service Studio structure
	///  <code>DataSeries</code> <p> Description: Data series to be used in graphs, the n cell of the valu
	/// e range will correspond to the n cell of the label range</p>
	/// </summary>
	[Serializable()]
	public partial struct STDataSeriesStructure: ISerializable, ITypedRecord<STDataSeriesStructure>, ISimpleRecord {
		private static readonly GlobalObjectKey IdValueRange = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*z6UPL__K0EKq_2sixzDIRQ");
		private static readonly GlobalObjectKey IdLabelRange = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*MtbosIZzqUKkocL3d4Zzhw");
		private static readonly GlobalObjectKey IdName = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*WqyHiIzy50W0NRqZbpcBKg");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("ValueRange")]
		public RCRangeRecord ssValueRange;

		[System.Xml.Serialization.XmlElement("LabelRange")]
		public RCRangeRecord ssLabelRange;

		[System.Xml.Serialization.XmlElement("Name")]
		public string ssName;


		public BitArray OptimizedAttributes;

		public STDataSeriesStructure(params string[] dummy) {
			OptimizedAttributes = null;
			ssValueRange = new RCRangeRecord(null);
			ssLabelRange = new RCRangeRecord(null);
			ssName = "";
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[2];
			all[0] = null;
			all[1] = null;
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
					ssValueRange.OptimizedAttributes = value[0];
					ssLabelRange.OptimizedAttributes = value[1];
				}
			}
			get {
				BitArray[] all = new BitArray[2];
				all[0] = null;
				all[1] = null;
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssName = r.ReadText(index++, "DataSeries.Name", "");
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
		public void ReadIM(STDataSeriesStructure r) {
			this = r;
		}


		public static bool operator == (STDataSeriesStructure a, STDataSeriesStructure b) {
			if (a.ssValueRange != b.ssValueRange) return false;
			if (a.ssLabelRange != b.ssLabelRange) return false;
			if (a.ssName != b.ssName) return false;
			return true;
		}

		public static bool operator != (STDataSeriesStructure a, STDataSeriesStructure b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(STDataSeriesStructure)) return false;
			return (this == (STDataSeriesStructure) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssValueRange.GetHashCode()
				^ ssLabelRange.GetHashCode()
				^ ssName.GetHashCode()
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

		public STDataSeriesStructure(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssValueRange = new RCRangeRecord(null);
			ssLabelRange = new RCRangeRecord(null);
			ssName = "";
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssValueRange", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssValueRange' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssValueRange = (RCRangeRecord) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssLabelRange", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssLabelRange' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssLabelRange = (RCRangeRecord) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssName", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssName' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssName = (string) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssValueRange.RecursiveReset();
			ssLabelRange.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssValueRange.InternalRecursiveSave();
			ssLabelRange.InternalRecursiveSave();
		}


		public STDataSeriesStructure Duplicate() {
			STDataSeriesStructure t;
			t.ssValueRange = (RCRangeRecord) this.ssValueRange.Duplicate();
			t.ssLabelRange = (RCRangeRecord) this.ssLabelRange.Duplicate();
			t.ssName = this.ssName;
			t.OptimizedAttributes = null;
			return t;
		}

		IRecord IRecord.Duplicate() {
			return Duplicate();
		}

		public void ToXml(Object parent, System.Xml.XmlElement baseElem, String fieldName, int detailLevel) {
			System.Xml.XmlElement recordElem = VarValue.AppendChild(baseElem, "Structure");
			if (fieldName != null) {
				VarValue.AppendAttribute(recordElem, "debug.field", fieldName);
				fieldName = fieldName.ToLowerInvariant();
			}
			if (detailLevel > 0) {
				ssValueRange.ToXml(this, recordElem, "ValueRange", detailLevel - 1);
				ssLabelRange.ToXml(this, recordElem, "LabelRange", detailLevel - 1);
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Name")) VarValue.AppendAttribute(recordElem, "Name", ssName, detailLevel, TypeKind.Text); else VarValue.AppendOptimizedAttribute(recordElem, "Name");
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "valuerange") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".ValueRange")) variable.Value = ssValueRange; else variable.Optimized = true;
				variable.SetFieldName("valuerange");
			} else if (head == "labelrange") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".LabelRange")) variable.Value = ssLabelRange; else variable.Optimized = true;
				variable.SetFieldName("labelrange");
			} else if (head == "name") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Name")) variable.Value = ssName; else variable.Optimized = true;
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
			if (key == IdValueRange) {
				return ssValueRange;
			} else if (key == IdLabelRange) {
				return ssLabelRange;
			} else if (key == IdName) {
				return ssName;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssValueRange.FillFromOther((IRecord) other.AttributeGet(IdValueRange));
			ssLabelRange.FillFromOther((IRecord) other.AttributeGet(IdLabelRange));
			ssName = (string) other.AttributeGet(IdName);
		}
		public bool IsDefault() {
			STDataSeriesStructure defaultStruct = new STDataSeriesStructure(null);
			if (this.ssValueRange != defaultStruct.ssValueRange) return false;
			if (this.ssLabelRange != defaultStruct.ssLabelRange) return false;
			if (this.ssName != defaultStruct.ssName) return false;
			return true;
		}
	} // STDataSeriesStructure

} // OutSystems.NssAdvanced_Excel
