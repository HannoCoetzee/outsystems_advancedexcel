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
		internal static readonly GlobalObjectKey IdFontName = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*BycYa7Www0ikxqPdTRgbGw");
		internal static readonly GlobalObjectKey IdFontSize = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*fQrThGBwgUybjIyLimeOPA");
		internal static readonly GlobalObjectKey IdBackgroundColor = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*3oXRIbtWvkKxhz4Cx7wIRg");
		internal static readonly GlobalObjectKey IdAutofitColumn = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*fH82ZL_ULky72619BBkR+w");
		internal static readonly GlobalObjectKey IdBold = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*GU0F2kjUe0W09TXtoXF07g");
		internal static readonly GlobalObjectKey IdNumberFormat = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*T9IMXF6dHkeMjRsrROBwig");
		internal static readonly GlobalObjectKey IdBorderStyle = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*4nZdbtDhvk+RtDYUBpdevA");
		internal static readonly GlobalObjectKey IdBorderColor = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*ZTYcPbwyLEuVyw+rm2CbRg");
		internal static readonly GlobalObjectKey IdHorizontalAlignment = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*VTZwPSIcVkmeFW4OW0ZguQ");
		internal static readonly GlobalObjectKey IdVerticalAlignment = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*ahrSd0mLV0Gc9jtbxoriRA");
		internal static readonly GlobalObjectKey IdWrapText = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*3uhZUV2hKkqHaaRD9xc48A");
		internal static readonly GlobalObjectKey IdShrinkToFit = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*paq0xKQFCEeqKMik8kiaNw");
		internal static readonly GlobalObjectKey IdTextRotation = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*XqOgxCqhiEyOcJ_Rg+aBsg");
		internal static readonly GlobalObjectKey IdHidden = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*vZcW2jua7E+q4utgGDRzeg");
		internal static readonly GlobalObjectKey IdIndent = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*JGWbJCDA00i7b+L3di8BPg");
		internal static readonly GlobalObjectKey IdLocked = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*OKB5o_cX0U+fvFVGpM3HQw");
		internal static readonly GlobalObjectKey IdQuotePrefix = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*FwRaU5uVeESO_8noGV00xw");
		internal static readonly GlobalObjectKey IdReadingOrder = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*xvpo+OQXa0uQmPDs2ZkhPQ");
		internal static readonly GlobalObjectKey IdFontColor = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*q3k5T2Uv00a1pJF4yVllTg");

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

		[System.Xml.Serialization.XmlElement("HorizontalAlignment")]
		public int ssHorizontalAlignment;

		[System.Xml.Serialization.XmlElement("VerticalAlignment")]
		public int ssVerticalAlignment;

		[System.Xml.Serialization.XmlElement("WrapText")]
		public bool ssWrapText;

		[System.Xml.Serialization.XmlElement("ShrinkToFit")]
		public bool ssShrinkToFit;

		[System.Xml.Serialization.XmlElement("TextRotation")]
		public int ssTextRotation;

		[System.Xml.Serialization.XmlElement("Hidden")]
		public bool ssHidden;

		[System.Xml.Serialization.XmlElement("Indent")]
		public int ssIndent;

		[System.Xml.Serialization.XmlElement("Locked")]
		public bool ssLocked;

		[System.Xml.Serialization.XmlElement("QuotePrefix")]
		public bool ssQuotePrefix;

		[System.Xml.Serialization.XmlElement("ReadingOrder")]
		public int ssReadingOrder;

		[System.Xml.Serialization.XmlElement("FontColor")]
		public string ssFontColor;


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
			ssHorizontalAlignment = 0;
			ssVerticalAlignment = 0;
			ssWrapText = false;
			ssShrinkToFit = false;
			ssTextRotation = 0;
			ssHidden = false;
			ssIndent = 0;
			ssLocked = true;
			ssQuotePrefix = false;
			ssReadingOrder = 0;
			ssFontColor = "";
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
			ssHorizontalAlignment = r.ReadInteger(index++, "CellFormat.HorizontalAlignment", 0);
			ssVerticalAlignment = r.ReadInteger(index++, "CellFormat.VerticalAlignment", 0);
			ssWrapText = r.ReadBoolean(index++, "CellFormat.WrapText", false);
			ssShrinkToFit = r.ReadBoolean(index++, "CellFormat.ShrinkToFit", false);
			ssTextRotation = r.ReadInteger(index++, "CellFormat.TextRotation", 0);
			ssHidden = r.ReadBoolean(index++, "CellFormat.Hidden", false);
			ssIndent = r.ReadInteger(index++, "CellFormat.Indent", 0);
			ssLocked = r.ReadBoolean(index++, "CellFormat.Locked", false);
			ssQuotePrefix = r.ReadBoolean(index++, "CellFormat.QuotePrefix", false);
			ssReadingOrder = r.ReadInteger(index++, "CellFormat.ReadingOrder", 0);
			ssFontColor = r.ReadText(index++, "CellFormat.FontColor", "");
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
			if (a.ssHorizontalAlignment != b.ssHorizontalAlignment) return false;
			if (a.ssVerticalAlignment != b.ssVerticalAlignment) return false;
			if (a.ssWrapText != b.ssWrapText) return false;
			if (a.ssShrinkToFit != b.ssShrinkToFit) return false;
			if (a.ssTextRotation != b.ssTextRotation) return false;
			if (a.ssHidden != b.ssHidden) return false;
			if (a.ssIndent != b.ssIndent) return false;
			if (a.ssLocked != b.ssLocked) return false;
			if (a.ssQuotePrefix != b.ssQuotePrefix) return false;
			if (a.ssReadingOrder != b.ssReadingOrder) return false;
			if (a.ssFontColor != b.ssFontColor) return false;
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
				^ ssHorizontalAlignment.GetHashCode()
				^ ssVerticalAlignment.GetHashCode()
				^ ssWrapText.GetHashCode()
				^ ssShrinkToFit.GetHashCode()
				^ ssTextRotation.GetHashCode()
				^ ssHidden.GetHashCode()
				^ ssIndent.GetHashCode()
				^ ssLocked.GetHashCode()
				^ ssQuotePrefix.GetHashCode()
				^ ssReadingOrder.GetHashCode()
				^ ssFontColor.GetHashCode()
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
			ssHorizontalAlignment = 0;
			ssVerticalAlignment = 0;
			ssWrapText = false;
			ssShrinkToFit = false;
			ssTextRotation = 0;
			ssHidden = false;
			ssIndent = 0;
			ssLocked = true;
			ssQuotePrefix = false;
			ssReadingOrder = 0;
			ssFontColor = "";
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
			fieldInfo = objInfo.GetField("ssHorizontalAlignment", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssHorizontalAlignment' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssHorizontalAlignment = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssVerticalAlignment", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssVerticalAlignment' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssVerticalAlignment = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssWrapText", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssWrapText' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssWrapText = (bool) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssShrinkToFit", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssShrinkToFit' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssShrinkToFit = (bool) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssTextRotation", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssTextRotation' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssTextRotation = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssHidden", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssHidden' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssHidden = (bool) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssIndent", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssIndent' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssIndent = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssLocked", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssLocked' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssLocked = (bool) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssQuotePrefix", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssQuotePrefix' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssQuotePrefix = (bool) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssReadingOrder", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssReadingOrder' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssReadingOrder = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssFontColor", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssFontColor' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssFontColor = (string) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
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
			t.ssHorizontalAlignment = this.ssHorizontalAlignment;
			t.ssVerticalAlignment = this.ssVerticalAlignment;
			t.ssWrapText = this.ssWrapText;
			t.ssShrinkToFit = this.ssShrinkToFit;
			t.ssTextRotation = this.ssTextRotation;
			t.ssHidden = this.ssHidden;
			t.ssIndent = this.ssIndent;
			t.ssLocked = this.ssLocked;
			t.ssQuotePrefix = this.ssQuotePrefix;
			t.ssReadingOrder = this.ssReadingOrder;
			t.ssFontColor = this.ssFontColor;
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
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".HorizontalAlignment")) VarValue.AppendAttribute(recordElem, "HorizontalAlignment", ssHorizontalAlignment, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "HorizontalAlignment");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".VerticalAlignment")) VarValue.AppendAttribute(recordElem, "VerticalAlignment", ssVerticalAlignment, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "VerticalAlignment");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".WrapText")) VarValue.AppendAttribute(recordElem, "WrapText", ssWrapText, detailLevel, TypeKind.Boolean); else VarValue.AppendOptimizedAttribute(recordElem, "WrapText");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".ShrinkToFit")) VarValue.AppendAttribute(recordElem, "ShrinkToFit", ssShrinkToFit, detailLevel, TypeKind.Boolean); else VarValue.AppendOptimizedAttribute(recordElem, "ShrinkToFit");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".TextRotation")) VarValue.AppendAttribute(recordElem, "TextRotation", ssTextRotation, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "TextRotation");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Hidden")) VarValue.AppendAttribute(recordElem, "Hidden", ssHidden, detailLevel, TypeKind.Boolean); else VarValue.AppendOptimizedAttribute(recordElem, "Hidden");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Indent")) VarValue.AppendAttribute(recordElem, "Indent", ssIndent, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "Indent");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Locked")) VarValue.AppendAttribute(recordElem, "Locked", ssLocked, detailLevel, TypeKind.Boolean); else VarValue.AppendOptimizedAttribute(recordElem, "Locked");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".QuotePrefix")) VarValue.AppendAttribute(recordElem, "QuotePrefix", ssQuotePrefix, detailLevel, TypeKind.Boolean); else VarValue.AppendOptimizedAttribute(recordElem, "QuotePrefix");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".ReadingOrder")) VarValue.AppendAttribute(recordElem, "ReadingOrder", ssReadingOrder, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "ReadingOrder");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".FontColor")) VarValue.AppendAttribute(recordElem, "FontColor", ssFontColor, detailLevel, TypeKind.Text); else VarValue.AppendOptimizedAttribute(recordElem, "FontColor");
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
			} else if (head == "horizontalalignment") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".HorizontalAlignment")) variable.Value = ssHorizontalAlignment; else variable.Optimized = true;
			} else if (head == "verticalalignment") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".VerticalAlignment")) variable.Value = ssVerticalAlignment; else variable.Optimized = true;
			} else if (head == "wraptext") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".WrapText")) variable.Value = ssWrapText; else variable.Optimized = true;
			} else if (head == "shrinktofit") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".ShrinkToFit")) variable.Value = ssShrinkToFit; else variable.Optimized = true;
			} else if (head == "textrotation") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".TextRotation")) variable.Value = ssTextRotation; else variable.Optimized = true;
			} else if (head == "hidden") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Hidden")) variable.Value = ssHidden; else variable.Optimized = true;
			} else if (head == "indent") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Indent")) variable.Value = ssIndent; else variable.Optimized = true;
			} else if (head == "locked") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Locked")) variable.Value = ssLocked; else variable.Optimized = true;
			} else if (head == "quoteprefix") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".QuotePrefix")) variable.Value = ssQuotePrefix; else variable.Optimized = true;
			} else if (head == "readingorder") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".ReadingOrder")) variable.Value = ssReadingOrder; else variable.Optimized = true;
			} else if (head == "fontcolor") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".FontColor")) variable.Value = ssFontColor; else variable.Optimized = true;
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
			} else if (key == IdHorizontalAlignment) {
				return ssHorizontalAlignment;
			} else if (key == IdVerticalAlignment) {
				return ssVerticalAlignment;
			} else if (key == IdWrapText) {
				return ssWrapText;
			} else if (key == IdShrinkToFit) {
				return ssShrinkToFit;
			} else if (key == IdTextRotation) {
				return ssTextRotation;
			} else if (key == IdHidden) {
				return ssHidden;
			} else if (key == IdIndent) {
				return ssIndent;
			} else if (key == IdLocked) {
				return ssLocked;
			} else if (key == IdQuotePrefix) {
				return ssQuotePrefix;
			} else if (key == IdReadingOrder) {
				return ssReadingOrder;
			} else if (key == IdFontColor) {
				return ssFontColor;
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
			ssHorizontalAlignment = (int) other.AttributeGet(IdHorizontalAlignment);
			ssVerticalAlignment = (int) other.AttributeGet(IdVerticalAlignment);
			ssWrapText = (bool) other.AttributeGet(IdWrapText);
			ssShrinkToFit = (bool) other.AttributeGet(IdShrinkToFit);
			ssTextRotation = (int) other.AttributeGet(IdTextRotation);
			ssHidden = (bool) other.AttributeGet(IdHidden);
			ssIndent = (int) other.AttributeGet(IdIndent);
			ssLocked = (bool) other.AttributeGet(IdLocked);
			ssQuotePrefix = (bool) other.AttributeGet(IdQuotePrefix);
			ssReadingOrder = (int) other.AttributeGet(IdReadingOrder);
			ssFontColor = (string) other.AttributeGet(IdFontColor);
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
			if (this.ssHorizontalAlignment != defaultStruct.ssHorizontalAlignment) return false;
			if (this.ssVerticalAlignment != defaultStruct.ssVerticalAlignment) return false;
			if (this.ssWrapText != defaultStruct.ssWrapText) return false;
			if (this.ssShrinkToFit != defaultStruct.ssShrinkToFit) return false;
			if (this.ssTextRotation != defaultStruct.ssTextRotation) return false;
			if (this.ssHidden != defaultStruct.ssHidden) return false;
			if (this.ssIndent != defaultStruct.ssIndent) return false;
			if (this.ssLocked != defaultStruct.ssLocked) return false;
			if (this.ssQuotePrefix != defaultStruct.ssQuotePrefix) return false;
			if (this.ssReadingOrder != defaultStruct.ssReadingOrder) return false;
			if (this.ssFontColor != defaultStruct.ssFontColor) return false;
			return true;
		}
	} // STCellFormatStructure

	/// <summary>
	/// Structure <code>STWorkbookStructure</code> that represents the Service Studio structure
	///  <code>Workbook</code> <p> Description: The Excel File</p>
	/// </summary>
	[Serializable()]
	public partial struct STWorkbookStructure: ISerializable, ITypedRecord<STWorkbookStructure>, ISimpleRecord {
		internal static readonly GlobalObjectKey IdWorksheets = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*izrPaRKZ_EqpgPLXabozRQ");

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
		internal static readonly GlobalObjectKey IdName = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*qu+kXpI5mk2nEoopujaw1w");
		internal static readonly GlobalObjectKey IdIndex = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*y5APBAwZfkWoTqWCP01Pxg");
		internal static readonly GlobalObjectKey IdTabColor = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*vXO_uSpdREW_4CDsVlbFlQ");
		internal static readonly GlobalObjectKey IdDimension = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*xXgLP3ev_0mowvNYBfWMGg");

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
		internal static readonly GlobalObjectKey IdIsKnownColor = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*l5cu6CGHqkSOaN_PCVYsDA");
		internal static readonly GlobalObjectKey IdIsNamedColor = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*eekKwslJf0u5aYbXGPnflA");
		internal static readonly GlobalObjectKey IdIsSystemColor = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*OD+P0K4xt02YijR2RL18uw");
		internal static readonly GlobalObjectKey IdA = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*zUlWKdvonEGw1Rc65EoX6w");
		internal static readonly GlobalObjectKey IdR = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*ECHSWe_rkkuNSf04BI6Adg");
		internal static readonly GlobalObjectKey IdG = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*5YqhHvTzz0aEouuGyeT4Ww");
		internal static readonly GlobalObjectKey IdB = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*c2IM4kriBEiuF+KhK3PbSg");
		internal static readonly GlobalObjectKey IdName = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*eyCFq+evuk6c3ZgZ2JJS_g");

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
		internal static readonly GlobalObjectKey IdAddress = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*yHVFLdcUqE+X718XHPI+kw");
		internal static readonly GlobalObjectKey IdColumns = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*aNFKux7b1kWMQpn59TfPGQ");
		internal static readonly GlobalObjectKey IdRows = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*01Ula9rHSkKShXRia3Q7sQ");
		internal static readonly GlobalObjectKey IdStart = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*JaS35nzoPk2ZuTVFbuOf0Q");
		internal static readonly GlobalObjectKey IdEnd = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*Gh6izWtZ+UKMWJvgHvCGSA");

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
		internal static readonly GlobalObjectKey IdAddress = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*AVbMSVppE0+cJ1ha0PX_Tg");
		internal static readonly GlobalObjectKey IdColumn = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*2K8ZTo2MUkGj+bwxIpHHSg");
		internal static readonly GlobalObjectKey IdIsRef = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*q_rA6foNt02HS9grG81W9w");
		internal static readonly GlobalObjectKey IdRow = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*+nGPqzJReEerUJI3NyiEPg");

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
		internal static readonly GlobalObjectKey IdStartRow = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*kAiDbHLyfka+BBBpQVgzag");
		internal static readonly GlobalObjectKey IdStartCol = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*LjVCwwGKH02mZjrBadmzXg");
		internal static readonly GlobalObjectKey IdEndRow = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*oW8H1BHvyEqUQO2kj1eHeQ");
		internal static readonly GlobalObjectKey IdEndCol = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*FidsGMKZGUKVmU2u5C7LgA");

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
		internal static readonly GlobalObjectKey IdValueRange = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*z6UPL__K0EKq_2sixzDIRQ");
		internal static readonly GlobalObjectKey IdLabelRange = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*MtbosIZzqUKkocL3d4Zzhw");
		internal static readonly GlobalObjectKey IdName = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*WqyHiIzy50W0NRqZbpcBKg");

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

	/// <summary>
	/// Structure <code>STConditionalFormatItemStructure</code> that represents the Service Studio
	///  structure <code>ConditionalFormatItem</code> <p> Description: Represents a conditional formattin
	/// g item.</p>
	/// </summary>
	[Serializable()]
	public partial struct STConditionalFormatItemStructure: ISerializable, ITypedRecord<STConditionalFormatItemStructure>, ISimpleRecord {
		internal static readonly GlobalObjectKey IdAddress = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*9quUxhHXV06FIXUHzAo0tA");
		internal static readonly GlobalObjectKey IdPriority = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*ZiyfnTaZTUSqZ9KxgazsJA");
		internal static readonly GlobalObjectKey IdStopIfTrue = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*V+9zUV7vZE6ULG3ngGiKqw");
		internal static readonly GlobalObjectKey IdFormula = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*bvwikWuyc0CFIjUukGNxMw");
		internal static readonly GlobalObjectKey IdRuleType = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*kAiJT_QaJEaD40wTFCo+pw");
		internal static readonly GlobalObjectKey IdStyle = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*CVx5EPnKoka3xe7InYEb7Q");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("Address")]
		public RCAddressRecord ssAddress;

		[System.Xml.Serialization.XmlElement("Priority")]
		public int ssPriority;

		[System.Xml.Serialization.XmlElement("StopIfTrue")]
		public bool ssStopIfTrue;

		[System.Xml.Serialization.XmlElement("Formula")]
		public string ssFormula;

		[System.Xml.Serialization.XmlElement("RuleType")]
		public int ssRuleType;

		[System.Xml.Serialization.XmlElement("Style")]
		public RCConditionalFormatStyleRecord ssStyle;


		public BitArray OptimizedAttributes;

		public STConditionalFormatItemStructure(params string[] dummy) {
			OptimizedAttributes = null;
			ssAddress = new RCAddressRecord(null);
			ssPriority = 101;
			ssStopIfTrue = false;
			ssFormula = "";
			ssRuleType = 0;
			ssStyle = new RCConditionalFormatStyleRecord(null);
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
					ssAddress.OptimizedAttributes = value[0];
					ssStyle.OptimizedAttributes = value[1];
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
			ssPriority = r.ReadInteger(index++, "ConditionalFormatItem.Priority", 0);
			ssStopIfTrue = r.ReadBoolean(index++, "ConditionalFormatItem.StopIfTrue", false);
			ssFormula = r.ReadText(index++, "ConditionalFormatItem.Formula", "");
			ssRuleType = r.ReadInteger(index++, "ConditionalFormatItem.RuleType", 0);
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
		public void ReadIM(STConditionalFormatItemStructure r) {
			this = r;
		}


		public static bool operator == (STConditionalFormatItemStructure a, STConditionalFormatItemStructure b) {
			if (a.ssAddress != b.ssAddress) return false;
			if (a.ssPriority != b.ssPriority) return false;
			if (a.ssStopIfTrue != b.ssStopIfTrue) return false;
			if (a.ssFormula != b.ssFormula) return false;
			if (a.ssRuleType != b.ssRuleType) return false;
			if (a.ssStyle != b.ssStyle) return false;
			return true;
		}

		public static bool operator != (STConditionalFormatItemStructure a, STConditionalFormatItemStructure b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(STConditionalFormatItemStructure)) return false;
			return (this == (STConditionalFormatItemStructure) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssAddress.GetHashCode()
				^ ssPriority.GetHashCode()
				^ ssStopIfTrue.GetHashCode()
				^ ssFormula.GetHashCode()
				^ ssRuleType.GetHashCode()
				^ ssStyle.GetHashCode()
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

		public STConditionalFormatItemStructure(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssAddress = new RCAddressRecord(null);
			ssPriority = 101;
			ssStopIfTrue = false;
			ssFormula = "";
			ssRuleType = 0;
			ssStyle = new RCConditionalFormatStyleRecord(null);
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssAddress", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssAddress' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssAddress = (RCAddressRecord) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssPriority", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssPriority' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssPriority = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssStopIfTrue", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssStopIfTrue' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssStopIfTrue = (bool) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssFormula", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssFormula' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssFormula = (string) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssRuleType", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssRuleType' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssRuleType = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssStyle", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssStyle' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssStyle = (RCConditionalFormatStyleRecord) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssAddress.RecursiveReset();
			ssStyle.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssAddress.InternalRecursiveSave();
			ssStyle.InternalRecursiveSave();
		}


		public STConditionalFormatItemStructure Duplicate() {
			STConditionalFormatItemStructure t;
			t.ssAddress = (RCAddressRecord) this.ssAddress.Duplicate();
			t.ssPriority = this.ssPriority;
			t.ssStopIfTrue = this.ssStopIfTrue;
			t.ssFormula = this.ssFormula;
			t.ssRuleType = this.ssRuleType;
			t.ssStyle = (RCConditionalFormatStyleRecord) this.ssStyle.Duplicate();
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
				ssAddress.ToXml(this, recordElem, "Address", detailLevel - 1);
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Priority")) VarValue.AppendAttribute(recordElem, "Priority", ssPriority, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "Priority");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".StopIfTrue")) VarValue.AppendAttribute(recordElem, "StopIfTrue", ssStopIfTrue, detailLevel, TypeKind.Boolean); else VarValue.AppendOptimizedAttribute(recordElem, "StopIfTrue");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Formula")) VarValue.AppendAttribute(recordElem, "Formula", ssFormula, detailLevel, TypeKind.Text); else VarValue.AppendOptimizedAttribute(recordElem, "Formula");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".RuleType")) VarValue.AppendAttribute(recordElem, "RuleType", ssRuleType, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "RuleType");
				ssStyle.ToXml(this, recordElem, "Style", detailLevel - 1);
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
				variable.SetFieldName("address");
			} else if (head == "priority") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Priority")) variable.Value = ssPriority; else variable.Optimized = true;
			} else if (head == "stopiftrue") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".StopIfTrue")) variable.Value = ssStopIfTrue; else variable.Optimized = true;
			} else if (head == "formula") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Formula")) variable.Value = ssFormula; else variable.Optimized = true;
			} else if (head == "ruletype") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".RuleType")) variable.Value = ssRuleType; else variable.Optimized = true;
			} else if (head == "style") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Style")) variable.Value = ssStyle; else variable.Optimized = true;
				variable.SetFieldName("style");
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
			} else if (key == IdPriority) {
				return ssPriority;
			} else if (key == IdStopIfTrue) {
				return ssStopIfTrue;
			} else if (key == IdFormula) {
				return ssFormula;
			} else if (key == IdRuleType) {
				return ssRuleType;
			} else if (key == IdStyle) {
				return ssStyle;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssAddress.FillFromOther((IRecord) other.AttributeGet(IdAddress));
			ssPriority = (int) other.AttributeGet(IdPriority);
			ssStopIfTrue = (bool) other.AttributeGet(IdStopIfTrue);
			ssFormula = (string) other.AttributeGet(IdFormula);
			ssRuleType = (int) other.AttributeGet(IdRuleType);
			ssStyle.FillFromOther((IRecord) other.AttributeGet(IdStyle));
		}
		public bool IsDefault() {
			STConditionalFormatItemStructure defaultStruct = new STConditionalFormatItemStructure(null);
			if (this.ssAddress != defaultStruct.ssAddress) return false;
			if (this.ssPriority != defaultStruct.ssPriority) return false;
			if (this.ssStopIfTrue != defaultStruct.ssStopIfTrue) return false;
			if (this.ssFormula != defaultStruct.ssFormula) return false;
			if (this.ssRuleType != defaultStruct.ssRuleType) return false;
			if (this.ssStyle != defaultStruct.ssStyle) return false;
			return true;
		}
	} // STConditionalFormatItemStructure

	/// <summary>
	/// Structure <code>STConditionalFormatStyleStructure</code> that represents the Service Studio
	///  structure <code>ConditionalFormatStyle</code> <p> Description: Style (format) options fo
	/// r Conditional Formatting.</p>
	/// </summary>
	[Serializable()]
	public partial struct STConditionalFormatStyleStructure: ISerializable, ITypedRecord<STConditionalFormatStyleStructure>, ISimpleRecord {
		internal static readonly GlobalObjectKey IdFill = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*8Lre4DmEEUGF3nDQEdYz4w");
		internal static readonly GlobalObjectKey IdNumberFormat = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*ScnLCO8PME+pYUGxccUiVQ");
		internal static readonly GlobalObjectKey IdBorderTop = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*Z1IJbohI1kW1D19RhHAzLA");
		internal static readonly GlobalObjectKey IdBorderBottom = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*iE4ewOcDZkqpxqiarvwtHA");
		internal static readonly GlobalObjectKey IdBorderLeft = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*C3y1_2KaaUmS52Km4SrEcg");
		internal static readonly GlobalObjectKey IdBorderRight = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*l1YNOqYcbUWXa8VCdoyuRQ");
		internal static readonly GlobalObjectKey IdFont = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*lLGusVRpd0G1Hsch057P7g");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("Fill")]
		public RCFillStyleRecord ssFill;

		[System.Xml.Serialization.XmlElement("NumberFormat")]
		public string ssNumberFormat;

		[System.Xml.Serialization.XmlElement("BorderTop")]
		public RCBorderStyleRecord ssBorderTop;

		[System.Xml.Serialization.XmlElement("BorderBottom")]
		public RCBorderStyleRecord ssBorderBottom;

		[System.Xml.Serialization.XmlElement("BorderLeft")]
		public RCBorderStyleRecord ssBorderLeft;

		[System.Xml.Serialization.XmlElement("BorderRight")]
		public RCBorderStyleRecord ssBorderRight;

		[System.Xml.Serialization.XmlElement("Font")]
		public RCFontStyleRecord ssFont;


		public BitArray OptimizedAttributes;

		public STConditionalFormatStyleStructure(params string[] dummy) {
			OptimizedAttributes = null;
			ssFill = new RCFillStyleRecord(null);
			ssNumberFormat = "";
			ssBorderTop = new RCBorderStyleRecord(null);
			ssBorderBottom = new RCBorderStyleRecord(null);
			ssBorderLeft = new RCBorderStyleRecord(null);
			ssBorderRight = new RCBorderStyleRecord(null);
			ssFont = new RCFontStyleRecord(null);
		}

		public BitArray[] GetDefaultOptimizedValues() {
			BitArray[] all = new BitArray[6];
			all[0] = null;
			all[1] = null;
			all[2] = null;
			all[3] = null;
			all[4] = null;
			all[5] = null;
			return all;
		}

		public BitArray[] AllOptimizedAttributes {
			set {
				if (value == null) {
				} else {
					ssFill.OptimizedAttributes = value[0];
					ssBorderTop.OptimizedAttributes = value[1];
					ssBorderBottom.OptimizedAttributes = value[2];
					ssBorderLeft.OptimizedAttributes = value[3];
					ssBorderRight.OptimizedAttributes = value[4];
					ssFont.OptimizedAttributes = value[5];
				}
			}
			get {
				BitArray[] all = new BitArray[6];
				all[0] = null;
				all[1] = null;
				all[2] = null;
				all[3] = null;
				all[4] = null;
				all[5] = null;
				return all;
			}
		}

		/// <summary>
		/// Read a record from database
		/// </summary>
		/// <param name="r"> Data base reader</param>
		/// <param name="index"> index</param>
		public void Read(IDataReader r, ref int index) {
			ssNumberFormat = r.ReadText(index++, "ConditionalFormatStyle.NumberFormat", "");
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
		public void ReadIM(STConditionalFormatStyleStructure r) {
			this = r;
		}


		public static bool operator == (STConditionalFormatStyleStructure a, STConditionalFormatStyleStructure b) {
			if (a.ssFill != b.ssFill) return false;
			if (a.ssNumberFormat != b.ssNumberFormat) return false;
			if (a.ssBorderTop != b.ssBorderTop) return false;
			if (a.ssBorderBottom != b.ssBorderBottom) return false;
			if (a.ssBorderLeft != b.ssBorderLeft) return false;
			if (a.ssBorderRight != b.ssBorderRight) return false;
			if (a.ssFont != b.ssFont) return false;
			return true;
		}

		public static bool operator != (STConditionalFormatStyleStructure a, STConditionalFormatStyleStructure b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(STConditionalFormatStyleStructure)) return false;
			return (this == (STConditionalFormatStyleStructure) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssFill.GetHashCode()
				^ ssNumberFormat.GetHashCode()
				^ ssBorderTop.GetHashCode()
				^ ssBorderBottom.GetHashCode()
				^ ssBorderLeft.GetHashCode()
				^ ssBorderRight.GetHashCode()
				^ ssFont.GetHashCode()
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

		public STConditionalFormatStyleStructure(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssFill = new RCFillStyleRecord(null);
			ssNumberFormat = "";
			ssBorderTop = new RCBorderStyleRecord(null);
			ssBorderBottom = new RCBorderStyleRecord(null);
			ssBorderLeft = new RCBorderStyleRecord(null);
			ssBorderRight = new RCBorderStyleRecord(null);
			ssFont = new RCFontStyleRecord(null);
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssFill", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssFill' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssFill = (RCFillStyleRecord) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssNumberFormat", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssNumberFormat' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssNumberFormat = (string) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssBorderTop", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssBorderTop' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssBorderTop = (RCBorderStyleRecord) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssBorderBottom", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssBorderBottom' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssBorderBottom = (RCBorderStyleRecord) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssBorderLeft", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssBorderLeft' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssBorderLeft = (RCBorderStyleRecord) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssBorderRight", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssBorderRight' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssBorderRight = (RCBorderStyleRecord) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssFont", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssFont' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssFont = (RCFontStyleRecord) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
			ssFill.RecursiveReset();
			ssBorderTop.RecursiveReset();
			ssBorderBottom.RecursiveReset();
			ssBorderLeft.RecursiveReset();
			ssBorderRight.RecursiveReset();
			ssFont.RecursiveReset();
		}

		public void InternalRecursiveSave() {
			ssFill.InternalRecursiveSave();
			ssBorderTop.InternalRecursiveSave();
			ssBorderBottom.InternalRecursiveSave();
			ssBorderLeft.InternalRecursiveSave();
			ssBorderRight.InternalRecursiveSave();
			ssFont.InternalRecursiveSave();
		}


		public STConditionalFormatStyleStructure Duplicate() {
			STConditionalFormatStyleStructure t;
			t.ssFill = (RCFillStyleRecord) this.ssFill.Duplicate();
			t.ssNumberFormat = this.ssNumberFormat;
			t.ssBorderTop = (RCBorderStyleRecord) this.ssBorderTop.Duplicate();
			t.ssBorderBottom = (RCBorderStyleRecord) this.ssBorderBottom.Duplicate();
			t.ssBorderLeft = (RCBorderStyleRecord) this.ssBorderLeft.Duplicate();
			t.ssBorderRight = (RCBorderStyleRecord) this.ssBorderRight.Duplicate();
			t.ssFont = (RCFontStyleRecord) this.ssFont.Duplicate();
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
				ssFill.ToXml(this, recordElem, "Fill", detailLevel - 1);
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".NumberFormat")) VarValue.AppendAttribute(recordElem, "NumberFormat", ssNumberFormat, detailLevel, TypeKind.Text); else VarValue.AppendOptimizedAttribute(recordElem, "NumberFormat");
				ssBorderTop.ToXml(this, recordElem, "BorderTop", detailLevel - 1);
				ssBorderBottom.ToXml(this, recordElem, "BorderBottom", detailLevel - 1);
				ssBorderLeft.ToXml(this, recordElem, "BorderLeft", detailLevel - 1);
				ssBorderRight.ToXml(this, recordElem, "BorderRight", detailLevel - 1);
				ssFont.ToXml(this, recordElem, "Font", detailLevel - 1);
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "fill") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Fill")) variable.Value = ssFill; else variable.Optimized = true;
				variable.SetFieldName("fill");
			} else if (head == "numberformat") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".NumberFormat")) variable.Value = ssNumberFormat; else variable.Optimized = true;
			} else if (head == "bordertop") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".BorderTop")) variable.Value = ssBorderTop; else variable.Optimized = true;
				variable.SetFieldName("bordertop");
			} else if (head == "borderbottom") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".BorderBottom")) variable.Value = ssBorderBottom; else variable.Optimized = true;
				variable.SetFieldName("borderbottom");
			} else if (head == "borderleft") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".BorderLeft")) variable.Value = ssBorderLeft; else variable.Optimized = true;
				variable.SetFieldName("borderleft");
			} else if (head == "borderright") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".BorderRight")) variable.Value = ssBorderRight; else variable.Optimized = true;
				variable.SetFieldName("borderright");
			} else if (head == "font") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Font")) variable.Value = ssFont; else variable.Optimized = true;
				variable.SetFieldName("font");
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
			if (key == IdFill) {
				return ssFill;
			} else if (key == IdNumberFormat) {
				return ssNumberFormat;
			} else if (key == IdBorderTop) {
				return ssBorderTop;
			} else if (key == IdBorderBottom) {
				return ssBorderBottom;
			} else if (key == IdBorderLeft) {
				return ssBorderLeft;
			} else if (key == IdBorderRight) {
				return ssBorderRight;
			} else if (key == IdFont) {
				return ssFont;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssFill.FillFromOther((IRecord) other.AttributeGet(IdFill));
			ssNumberFormat = (string) other.AttributeGet(IdNumberFormat);
			ssBorderTop.FillFromOther((IRecord) other.AttributeGet(IdBorderTop));
			ssBorderBottom.FillFromOther((IRecord) other.AttributeGet(IdBorderBottom));
			ssBorderLeft.FillFromOther((IRecord) other.AttributeGet(IdBorderLeft));
			ssBorderRight.FillFromOther((IRecord) other.AttributeGet(IdBorderRight));
			ssFont.FillFromOther((IRecord) other.AttributeGet(IdFont));
		}
		public bool IsDefault() {
			STConditionalFormatStyleStructure defaultStruct = new STConditionalFormatStyleStructure(null);
			if (this.ssFill != defaultStruct.ssFill) return false;
			if (this.ssNumberFormat != defaultStruct.ssNumberFormat) return false;
			if (this.ssBorderTop != defaultStruct.ssBorderTop) return false;
			if (this.ssBorderBottom != defaultStruct.ssBorderBottom) return false;
			if (this.ssBorderLeft != defaultStruct.ssBorderLeft) return false;
			if (this.ssBorderRight != defaultStruct.ssBorderRight) return false;
			if (this.ssFont != defaultStruct.ssFont) return false;
			return true;
		}
	} // STConditionalFormatStyleStructure

	/// <summary>
	/// Structure <code>STBorderStyleStructure</code> that represents the Service Studio structure
	///  <code>BorderStyle</code> <p> Description: Style and color of border</p>
	/// </summary>
	[Serializable()]
	public partial struct STBorderStyleStructure: ISerializable, ITypedRecord<STBorderStyleStructure>, ISimpleRecord {
		internal static readonly GlobalObjectKey IdStyle = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*j+Bh82+j9kGFWhNQ0zRlEg");
		internal static readonly GlobalObjectKey IdColor = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*lJE4e2V5rUWa3RMyGg_CqA");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("Style")]
		public int ssStyle;

		[System.Xml.Serialization.XmlElement("Color")]
		public string ssColor;


		public BitArray OptimizedAttributes;

		public STBorderStyleStructure(params string[] dummy) {
			OptimizedAttributes = null;
			ssStyle = 0;
			ssColor = "";
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
			ssStyle = r.ReadInteger(index++, "BorderStyle.Style", 0);
			ssColor = r.ReadText(index++, "BorderStyle.Color", "");
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
		public void ReadIM(STBorderStyleStructure r) {
			this = r;
		}


		public static bool operator == (STBorderStyleStructure a, STBorderStyleStructure b) {
			if (a.ssStyle != b.ssStyle) return false;
			if (a.ssColor != b.ssColor) return false;
			return true;
		}

		public static bool operator != (STBorderStyleStructure a, STBorderStyleStructure b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(STBorderStyleStructure)) return false;
			return (this == (STBorderStyleStructure) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssStyle.GetHashCode()
				^ ssColor.GetHashCode()
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

		public STBorderStyleStructure(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssStyle = 0;
			ssColor = "";
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssStyle", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssStyle' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssStyle = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssColor", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssColor' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssColor = (string) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
		}

		public void InternalRecursiveSave() {
		}


		public STBorderStyleStructure Duplicate() {
			STBorderStyleStructure t;
			t.ssStyle = this.ssStyle;
			t.ssColor = this.ssColor;
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
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Style")) VarValue.AppendAttribute(recordElem, "Style", ssStyle, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "Style");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Color")) VarValue.AppendAttribute(recordElem, "Color", ssColor, detailLevel, TypeKind.Text); else VarValue.AppendOptimizedAttribute(recordElem, "Color");
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "style") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Style")) variable.Value = ssStyle; else variable.Optimized = true;
			} else if (head == "color") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Color")) variable.Value = ssColor; else variable.Optimized = true;
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
			if (key == IdStyle) {
				return ssStyle;
			} else if (key == IdColor) {
				return ssColor;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssStyle = (int) other.AttributeGet(IdStyle);
			ssColor = (string) other.AttributeGet(IdColor);
		}
		public bool IsDefault() {
			STBorderStyleStructure defaultStruct = new STBorderStyleStructure(null);
			if (this.ssStyle != defaultStruct.ssStyle) return false;
			if (this.ssColor != defaultStruct.ssColor) return false;
			return true;
		}
	} // STBorderStyleStructure

	/// <summary>
	/// Structure <code>STFillStyleStructure</code> that represents the Service Studio structure
	///  <code>FillStyle</code> <p> Description: Fill properties</p>
	/// </summary>
	[Serializable()]
	public partial struct STFillStyleStructure: ISerializable, ITypedRecord<STFillStyleStructure>, ISimpleRecord {
		internal static readonly GlobalObjectKey IdBackgroundColor = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*PYSKmk8IJ0SWPNkGHZI+nw");
		internal static readonly GlobalObjectKey IdPatternColor = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*67yv3NhuPUudYNgC61tldQ");
		internal static readonly GlobalObjectKey IdPatternType = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*Wxly5fCSdEqAzKDA6dBrhA");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("BackgroundColor")]
		public string ssBackgroundColor;

		[System.Xml.Serialization.XmlElement("PatternColor")]
		public string ssPatternColor;

		[System.Xml.Serialization.XmlElement("PatternType")]
		public int ssPatternType;


		public BitArray OptimizedAttributes;

		public STFillStyleStructure(params string[] dummy) {
			OptimizedAttributes = null;
			ssBackgroundColor = "";
			ssPatternColor = "";
			ssPatternType = 0;
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
			ssBackgroundColor = r.ReadText(index++, "FillStyle.BackgroundColor", "");
			ssPatternColor = r.ReadText(index++, "FillStyle.PatternColor", "");
			ssPatternType = r.ReadInteger(index++, "FillStyle.PatternType", 0);
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
		public void ReadIM(STFillStyleStructure r) {
			this = r;
		}


		public static bool operator == (STFillStyleStructure a, STFillStyleStructure b) {
			if (a.ssBackgroundColor != b.ssBackgroundColor) return false;
			if (a.ssPatternColor != b.ssPatternColor) return false;
			if (a.ssPatternType != b.ssPatternType) return false;
			return true;
		}

		public static bool operator != (STFillStyleStructure a, STFillStyleStructure b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(STFillStyleStructure)) return false;
			return (this == (STFillStyleStructure) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssBackgroundColor.GetHashCode()
				^ ssPatternColor.GetHashCode()
				^ ssPatternType.GetHashCode()
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

		public STFillStyleStructure(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssBackgroundColor = "";
			ssPatternColor = "";
			ssPatternType = 0;
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssBackgroundColor", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssBackgroundColor' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssBackgroundColor = (string) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssPatternColor", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssPatternColor' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssPatternColor = (string) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssPatternType", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssPatternType' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssPatternType = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
		}

		public void InternalRecursiveSave() {
		}


		public STFillStyleStructure Duplicate() {
			STFillStyleStructure t;
			t.ssBackgroundColor = this.ssBackgroundColor;
			t.ssPatternColor = this.ssPatternColor;
			t.ssPatternType = this.ssPatternType;
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
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".BackgroundColor")) VarValue.AppendAttribute(recordElem, "BackgroundColor", ssBackgroundColor, detailLevel, TypeKind.Text); else VarValue.AppendOptimizedAttribute(recordElem, "BackgroundColor");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".PatternColor")) VarValue.AppendAttribute(recordElem, "PatternColor", ssPatternColor, detailLevel, TypeKind.Text); else VarValue.AppendOptimizedAttribute(recordElem, "PatternColor");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".PatternType")) VarValue.AppendAttribute(recordElem, "PatternType", ssPatternType, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "PatternType");
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "backgroundcolor") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".BackgroundColor")) variable.Value = ssBackgroundColor; else variable.Optimized = true;
			} else if (head == "patterncolor") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".PatternColor")) variable.Value = ssPatternColor; else variable.Optimized = true;
			} else if (head == "patterntype") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".PatternType")) variable.Value = ssPatternType; else variable.Optimized = true;
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
			if (key == IdBackgroundColor) {
				return ssBackgroundColor;
			} else if (key == IdPatternColor) {
				return ssPatternColor;
			} else if (key == IdPatternType) {
				return ssPatternType;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssBackgroundColor = (string) other.AttributeGet(IdBackgroundColor);
			ssPatternColor = (string) other.AttributeGet(IdPatternColor);
			ssPatternType = (int) other.AttributeGet(IdPatternType);
		}
		public bool IsDefault() {
			STFillStyleStructure defaultStruct = new STFillStyleStructure(null);
			if (this.ssBackgroundColor != defaultStruct.ssBackgroundColor) return false;
			if (this.ssPatternColor != defaultStruct.ssPatternColor) return false;
			if (this.ssPatternType != defaultStruct.ssPatternType) return false;
			return true;
		}
	} // STFillStyleStructure

	/// <summary>
	/// Structure <code>STFontStyleStructure</code> that represents the Service Studio structure
	///  <code>FontStyle</code> <p> Description: </p>
	/// </summary>
	[Serializable()]
	public partial struct STFontStyleStructure: ISerializable, ITypedRecord<STFontStyleStructure>, ISimpleRecord {
		internal static readonly GlobalObjectKey IdBold = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*8fsIq2UKukCxtqCm89cwKw");
		internal static readonly GlobalObjectKey IdItalic = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*_iYwYIFAR0WMCT_NxN9F1w");
		internal static readonly GlobalObjectKey IdStrike = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*uJNOSb7cQUScvCImDAYnHw");
		internal static readonly GlobalObjectKey IdColor = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*7O071h1R9UOTFc35XugTxA");
		internal static readonly GlobalObjectKey IdUnderline = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*0_1vqRNHpUSMHYrgwUD9uA");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("Bold")]
		public bool ssBold;

		[System.Xml.Serialization.XmlElement("Italic")]
		public bool ssItalic;

		[System.Xml.Serialization.XmlElement("Strike")]
		public bool ssStrike;

		[System.Xml.Serialization.XmlElement("Color")]
		public string ssColor;

		[System.Xml.Serialization.XmlElement("Underline")]
		public int ssUnderline;


		public BitArray OptimizedAttributes;

		public STFontStyleStructure(params string[] dummy) {
			OptimizedAttributes = null;
			ssBold = false;
			ssItalic = false;
			ssStrike = false;
			ssColor = "";
			ssUnderline = 0;
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
			ssBold = r.ReadBoolean(index++, "FontStyle.Bold", false);
			ssItalic = r.ReadBoolean(index++, "FontStyle.Italic", false);
			ssStrike = r.ReadBoolean(index++, "FontStyle.Strike", false);
			ssColor = r.ReadText(index++, "FontStyle.Color", "");
			ssUnderline = r.ReadInteger(index++, "FontStyle.Underline", 0);
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
		public void ReadIM(STFontStyleStructure r) {
			this = r;
		}


		public static bool operator == (STFontStyleStructure a, STFontStyleStructure b) {
			if (a.ssBold != b.ssBold) return false;
			if (a.ssItalic != b.ssItalic) return false;
			if (a.ssStrike != b.ssStrike) return false;
			if (a.ssColor != b.ssColor) return false;
			if (a.ssUnderline != b.ssUnderline) return false;
			return true;
		}

		public static bool operator != (STFontStyleStructure a, STFontStyleStructure b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(STFontStyleStructure)) return false;
			return (this == (STFontStyleStructure) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssBold.GetHashCode()
				^ ssItalic.GetHashCode()
				^ ssStrike.GetHashCode()
				^ ssColor.GetHashCode()
				^ ssUnderline.GetHashCode()
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

		public STFontStyleStructure(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssBold = false;
			ssItalic = false;
			ssStrike = false;
			ssColor = "";
			ssUnderline = 0;
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssBold", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssBold' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssBold = (bool) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssItalic", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssItalic' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssItalic = (bool) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssStrike", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssStrike' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssStrike = (bool) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssColor", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssColor' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssColor = (string) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssUnderline", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssUnderline' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssUnderline = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
		}

		public void InternalRecursiveSave() {
		}


		public STFontStyleStructure Duplicate() {
			STFontStyleStructure t;
			t.ssBold = this.ssBold;
			t.ssItalic = this.ssItalic;
			t.ssStrike = this.ssStrike;
			t.ssColor = this.ssColor;
			t.ssUnderline = this.ssUnderline;
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
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Bold")) VarValue.AppendAttribute(recordElem, "Bold", ssBold, detailLevel, TypeKind.Boolean); else VarValue.AppendOptimizedAttribute(recordElem, "Bold");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Italic")) VarValue.AppendAttribute(recordElem, "Italic", ssItalic, detailLevel, TypeKind.Boolean); else VarValue.AppendOptimizedAttribute(recordElem, "Italic");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Strike")) VarValue.AppendAttribute(recordElem, "Strike", ssStrike, detailLevel, TypeKind.Boolean); else VarValue.AppendOptimizedAttribute(recordElem, "Strike");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Color")) VarValue.AppendAttribute(recordElem, "Color", ssColor, detailLevel, TypeKind.Text); else VarValue.AppendOptimizedAttribute(recordElem, "Color");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Underline")) VarValue.AppendAttribute(recordElem, "Underline", ssUnderline, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "Underline");
			} else {
				VarValue.AppendDeferredEvaluationElement(recordElem);
			}
		}

		public void EvaluateFields(VarValue variable, Object parent, String baseName, String fields) {
			String head = VarValue.GetHead(fields);
			String tail = VarValue.GetTail(fields);
			variable.Found = false;
			if (head == "bold") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Bold")) variable.Value = ssBold; else variable.Optimized = true;
			} else if (head == "italic") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Italic")) variable.Value = ssItalic; else variable.Optimized = true;
			} else if (head == "strike") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Strike")) variable.Value = ssStrike; else variable.Optimized = true;
			} else if (head == "color") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Color")) variable.Value = ssColor; else variable.Optimized = true;
			} else if (head == "underline") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Underline")) variable.Value = ssUnderline; else variable.Optimized = true;
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
			if (key == IdBold) {
				return ssBold;
			} else if (key == IdItalic) {
				return ssItalic;
			} else if (key == IdStrike) {
				return ssStrike;
			} else if (key == IdColor) {
				return ssColor;
			} else if (key == IdUnderline) {
				return ssUnderline;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssBold = (bool) other.AttributeGet(IdBold);
			ssItalic = (bool) other.AttributeGet(IdItalic);
			ssStrike = (bool) other.AttributeGet(IdStrike);
			ssColor = (string) other.AttributeGet(IdColor);
			ssUnderline = (int) other.AttributeGet(IdUnderline);
		}
		public bool IsDefault() {
			STFontStyleStructure defaultStruct = new STFontStyleStructure(null);
			if (this.ssBold != defaultStruct.ssBold) return false;
			if (this.ssItalic != defaultStruct.ssItalic) return false;
			if (this.ssStrike != defaultStruct.ssStrike) return false;
			if (this.ssColor != defaultStruct.ssColor) return false;
			if (this.ssUnderline != defaultStruct.ssUnderline) return false;
			return true;
		}
	} // STFontStyleStructure

	/// <summary>
	/// Structure <code>STNewSheetStructure</code> that represents the Service Studio structure
	///  <code>NewSheet</code> <p> Description: New worksheet object</p>
	/// </summary>
	[Serializable()]
	public partial struct STNewSheetStructure: ISerializable, ITypedRecord<STNewSheetStructure>, ISimpleRecord {
		internal static readonly GlobalObjectKey IdName = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*T_pKIf3Ftk2QFFghtZPzTA");
		internal static readonly GlobalObjectKey IdIndex = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*Gd9zj1aJrUKHcrFpKwI6tA");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("Name")]
		public string ssName;

		[System.Xml.Serialization.XmlElement("Index")]
		public int ssIndex;


		public BitArray OptimizedAttributes;

		public STNewSheetStructure(params string[] dummy) {
			OptimizedAttributes = null;
			ssName = "";
			ssIndex = 0;
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
			ssName = r.ReadText(index++, "NewSheet.Name", "");
			ssIndex = r.ReadInteger(index++, "NewSheet.Index", 0);
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
		public void ReadIM(STNewSheetStructure r) {
			this = r;
		}


		public static bool operator == (STNewSheetStructure a, STNewSheetStructure b) {
			if (a.ssName != b.ssName) return false;
			if (a.ssIndex != b.ssIndex) return false;
			return true;
		}

		public static bool operator != (STNewSheetStructure a, STNewSheetStructure b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(STNewSheetStructure)) return false;
			return (this == (STNewSheetStructure) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssName.GetHashCode()
				^ ssIndex.GetHashCode()
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

		public STNewSheetStructure(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssName = "";
			ssIndex = 0;
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
		}

		public void RecursiveReset() {
		}

		public void InternalRecursiveSave() {
		}


		public STNewSheetStructure Duplicate() {
			STNewSheetStructure t;
			t.ssName = this.ssName;
			t.ssIndex = this.ssIndex;
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
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssName = (string) other.AttributeGet(IdName);
			ssIndex = (int) other.AttributeGet(IdIndex);
		}
		public bool IsDefault() {
			STNewSheetStructure defaultStruct = new STNewSheetStructure(null);
			if (this.ssName != defaultStruct.ssName) return false;
			if (this.ssIndex != defaultStruct.ssIndex) return false;
			return true;
		}
	} // STNewSheetStructure

	/// <summary>
	/// Structure <code>STImageStructure</code> that represents the Service Studio structure
	///  <code>Image</code> <p> Description: </p>
	/// </summary>
	[Serializable()]
	public partial struct STImageStructure: ISerializable, ITypedRecord<STImageStructure>, ISimpleRecord {
		internal static readonly GlobalObjectKey IdName = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*658By1ms2kWlOg_UYpTHXA");
		internal static readonly GlobalObjectKey IdContent = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*3no2SVVvxkO8zFBWeiCOTw");
		internal static readonly GlobalObjectKey IdRow = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*wXDmnydcA0iEGPTBG_EY2A");
		internal static readonly GlobalObjectKey IdColumn = GlobalObjectKey.Parse("tQrPfipdPE2fHQ34mD74Uw*kewa5eHGGEyXFYsIBtKYjA");

		public static void EnsureInitialized() {}
		[System.Xml.Serialization.XmlElement("Name")]
		public string ssName;

		[System.Xml.Serialization.XmlElement("Content")]
		public byte[] ssContent;

		[System.Xml.Serialization.XmlElement("Row")]
		public int ssRow;

		[System.Xml.Serialization.XmlElement("Column")]
		public int ssColumn;


		public BitArray OptimizedAttributes;

		public STImageStructure(params string[] dummy) {
			OptimizedAttributes = null;
			ssName = "";
			ssContent = new byte[] {};
			ssRow = 0;
			ssColumn = 0;
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
			ssName = r.ReadText(index++, "Image.Name", "");
			ssContent = r.ReadBinaryData(index++, "Image.Content", new byte[] {});
			ssRow = r.ReadInteger(index++, "Image.Row", 0);
			ssColumn = r.ReadInteger(index++, "Image.Column", 0);
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
		public void ReadIM(STImageStructure r) {
			this = r;
		}


		public static bool operator == (STImageStructure a, STImageStructure b) {
			if (a.ssName != b.ssName) return false;
			if (!RuntimePlatformUtils.CompareByteArrays(a.ssContent, b.ssContent)) return false;
			if (a.ssRow != b.ssRow) return false;
			if (a.ssColumn != b.ssColumn) return false;
			return true;
		}

		public static bool operator != (STImageStructure a, STImageStructure b) {
			return !(a==b);
		}

		public override bool Equals(object o) {
			if (o.GetType() != typeof(STImageStructure)) return false;
			return (this == (STImageStructure) o);
		}

		public override int GetHashCode() {
			try {
				return base.GetHashCode()
				^ ssName.GetHashCode()
				^ ssContent.GetHashCode()
				^ ssRow.GetHashCode()
				^ ssColumn.GetHashCode()
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

		public STImageStructure(SerializationInfo info, StreamingContext context) {
			OptimizedAttributes = null;
			ssName = "";
			ssContent = new byte[] {};
			ssRow = 0;
			ssColumn = 0;
			Type objInfo = this.GetType();
			FieldInfo fieldInfo = null;
			fieldInfo = objInfo.GetField("ssName", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssName' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssName = (string) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssContent", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssContent' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssContent = (byte[]) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssRow", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssRow' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssRow = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
			fieldInfo = objInfo.GetField("ssColumn", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);
			if (fieldInfo == null) {
				throw new Exception("The field named 'ssColumn' was not found.");
			}
			if (fieldInfo.FieldType.IsSerializable) {
				ssColumn = (int) info.GetValue(fieldInfo.Name, fieldInfo.FieldType);
			}
		}

		public void RecursiveReset() {
		}

		public void InternalRecursiveSave() {
		}


		public STImageStructure Duplicate() {
			STImageStructure t;
			t.ssName = this.ssName;
			if (this.ssContent != null) {
				t.ssContent = (byte[]) this.ssContent.Clone();
			} else {
				t.ssContent = null;
			}
			t.ssRow = this.ssRow;
			t.ssColumn = this.ssColumn;
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
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Content")) VarValue.AppendAttribute(recordElem, "Content", ssContent, detailLevel, TypeKind.BinaryData); else VarValue.AppendOptimizedAttribute(recordElem, "Content");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Row")) VarValue.AppendAttribute(recordElem, "Row", ssRow, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "Row");
				if (!VarValue.FieldIsOptimized(parent, fieldName + ".Column")) VarValue.AppendAttribute(recordElem, "Column", ssColumn, detailLevel, TypeKind.Integer); else VarValue.AppendOptimizedAttribute(recordElem, "Column");
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
			} else if (head == "content") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Content")) variable.Value = ssContent; else variable.Optimized = true;
			} else if (head == "row") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Row")) variable.Value = ssRow; else variable.Optimized = true;
			} else if (head == "column") {
				if (!VarValue.FieldIsOptimized(parent, baseName + ".Column")) variable.Value = ssColumn; else variable.Optimized = true;
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
			} else if (key == IdContent) {
				return ssContent;
			} else if (key == IdRow) {
				return ssRow;
			} else if (key == IdColumn) {
				return ssColumn;
			} else {
				throw new Exception("Invalid key");
			}
		}
		public void FillFromOther(IRecord other) {
			if (other == null) return;
			ssName = (string) other.AttributeGet(IdName);
			ssContent = (byte[]) other.AttributeGet(IdContent);
			ssRow = (int) other.AttributeGet(IdRow);
			ssColumn = (int) other.AttributeGet(IdColumn);
		}
		public bool IsDefault() {
			STImageStructure defaultStruct = new STImageStructure(null);
			if (this.ssName != defaultStruct.ssName) return false;
			if (!RuntimePlatformUtils.CompareByteArrays(this.ssContent, defaultStruct.ssContent)) return false;
			if (this.ssRow != defaultStruct.ssRow) return false;
			if (this.ssColumn != defaultStruct.ssColumn) return false;
			return true;
		}
	} // STImageStructure

} // OutSystems.NssAdvanced_Excel
