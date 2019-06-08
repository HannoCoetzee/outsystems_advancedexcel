using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using System.Data;
using System.Reflection;
using OutSystems.HubEdition.RuntimePlatform.Db;


namespace OutSystems.NssAdvanced_Excel
{
    class Util
    {
        public static DataTable ConvertArrayListToDataTable(IList<IRecord> arrayList)
        {
            DataTable dt = new DataTable();

            if (arrayList.Count != 0)
            {
                dt = ConvertObjectToDataTableSchema(arrayList[0]);


                FillData(arrayList, dt);
            }

            return dt;
        }

        public static DataTable ConvertObjectToDataTableSchema(Object o)
        {
            DataTable dt = new DataTable();
            // get all fields for given row 
            FieldInfo fieldInfo = o.GetType().GetFields()[0];
            foreach (FieldInfo field in fieldInfo.GetValue(o).GetType().GetFields()) // columns/fields            
            {
                DataColumn dc = new DataColumn(field.Name);
                dc.DataType = field.FieldType;
                dt.Columns.Add(dc);
            }
            return dt;
        }

        private static void FillData(IList<IRecord> arrayList, DataTable dt)
        {
            foreach (Object o in arrayList)
            {
                DataRow dr = dt.NewRow();
                FieldInfo fieldInfo = o.GetType().GetFields()[0];

                DateTime nullDate = new DateTime(1901, 01, 01);
                DateTime d = DateTime.MinValue;

                foreach (FieldInfo field in fieldInfo.GetValue(o).GetType().GetFields()) // columns/fields                
                {
                    if (field.FieldType == typeof(System.DateTime))
                    {
                        d = Convert.ToDateTime(field.GetValue(fieldInfo.GetValue(o)));
                        if (d.CompareTo(nullDate) >= 1) dr[field.Name] = field.GetValue(fieldInfo.GetValue(o));
                    }
                    else
                        dr[field.Name] = field.GetValue(fieldInfo.GetValue(o));
                }
                dt.Rows.Add(dr);
            }
        }
    }
}
