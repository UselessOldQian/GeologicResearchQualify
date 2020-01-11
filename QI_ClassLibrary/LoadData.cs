using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Xml.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Controls;

namespace QI_ClassLibrary
{
    public class LoadData
    {
        /// <summary>
        /// 将Itable数据显示在DataGridView中
        /// </summary>
        /// <param name="ptable">Itable</param>
        /// <param name="DGV">DataGridView</param>
        /// <param name="pCursor">Cursor</param>
        public DataTable ShowTableInDataGridView(ITable ptable, DataGridView DGV, ref ICursor pCursor, ref IRow pRrow, out List<String> FieldName)
        {
            DGV.DataSource = null;
            DataTable pDataTable = new DataTable();//建立一个table
            FieldName = new List<string>();
            for (int i = 0; i < ptable.Fields.FieldCount; i++)
            {
                //建立一个string变量存储Field的名字
                FieldName.Add(ptable.Fields.get_Field(i).AliasName);
                string FieldTrueName = ptable.Fields.get_Field(i).Name;
                pDataTable.Columns.Add(FieldTrueName);
            }
            int index = 0;
            pCursor = ptable.Search(null, false);
            pRrow = pCursor.NextRow();
            while (pRrow != null && index < 1000)
            {
                DataRow pRow = pDataTable.NewRow();
                string[] StrRow = new string[pRrow.Fields.FieldCount];
                for (int i = 0; i < pRrow.Fields.FieldCount; i++)
                {
                    StrRow[i] = pRrow.get_Value(i).ToString();
                }
                pRow.ItemArray = StrRow;
                pDataTable.Rows.Add(pRow);
                pRrow = pCursor.NextRow();
                index++;
            }
            DGV.DataSource = pDataTable;
            for(int i =0;i<FieldName.Count;i++)
            {
                //if (ptable.Fields.get_Field(i).Type == esriFieldType.esriFieldTypeDate) { DGV.Columns[i].ValueType = typeof.};
                DGV.Columns[i].HeaderText = FieldName[i];
            }
            return pDataTable;
        }

        /// <summary>
        /// 根据Cursor和Row获取下1000行的数据
        /// </summary>
        /// <param name="pCursor">Cursor</param>
        /// <param name="pRrow">Row</param>
        /// <returns>下1000行数据的DataTable</returns>
        public DataTable GetData(ref ICursor pCursor, ref IRow pRrow)
        {
            DataTable dt = new DataTable();//声明DataSet对象
            if (pRrow == null) { return null; }
            for (int i = 0; i < pRrow.Fields.FieldCount; i++)
            {
                string FieldName;//建立一个string变量存储Field的名字
                FieldName = pRrow.Fields.get_Field(i).AliasName;
                dt.Columns.Add(FieldName);
            }
            int index = 0;
            while (pRrow != null && index < 1000)
            {
                DataRow pRow = dt.NewRow();
                string[] StrRow = new string[pRrow.Fields.FieldCount];
                for (int i = 0; i < pRrow.Fields.FieldCount; i++)
                {
                    StrRow[i] = pRrow.get_Value(i).ToString();
                }
                pRow.ItemArray = StrRow;
                dt.Rows.Add(pRow);
                pRrow = pCursor.NextRow();
                index++;
            }
            return dt;//返回
        }

        public DataTable ShowTableInDataGridView_zenjian(ITable ptable, DataGridView DGV,out List<String> FieldName)
        {
            DGV.DataSource = null;
            DataTable pDataTable = new DataTable();//建立一个table
            FieldName = new List<string>();
            for (int i = 0; i < ptable.Fields.FieldCount; i++)
            {
                //建立一个string变量存储Field的名字
                FieldName.Add(ptable.Fields.get_Field(i).AliasName);
                string FieldTrueName = ptable.Fields.get_Field(i).Name;
                pDataTable.Columns.Add(FieldTrueName);
            }
            int index = 0;
            ICursor pCursor = ptable.Search(null, false);
            IRow pRrow = pCursor.NextRow();
            while (pRrow != null)
            {
                DataRow pRow = pDataTable.NewRow();
                string[] StrRow = new string[pRrow.Fields.FieldCount];
                for (int i = 0; i < pRrow.Fields.FieldCount; i++)
                {
                    StrRow[i] = pRrow.get_Value(i).ToString();
                }
                pRow.ItemArray = StrRow;
                pDataTable.Rows.Add(pRow);
                pRrow = pCursor.NextRow();
                index++;
            }
            DGV.DataSource = pDataTable;
            for (int i = 0; i < FieldName.Count; i++)
            {
                //if (ptable.Fields.get_Field(i).Type == esriFieldType.esriFieldTypeDate) { DGV.Columns[i].ValueType = typeof.};
                DGV.Columns[i].HeaderText = FieldName[i];
            }
            return pDataTable;
        }

    }
}
