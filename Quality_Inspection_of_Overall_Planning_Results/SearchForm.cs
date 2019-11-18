using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Geodatabase;

namespace Quality_Inspection_of_Overall_Planning_Results
{
    public partial class SearchForm : Form
    {
        public event EventHandler<SQLFileterEventArgs> SqlOK;
        private List<ILayer> _layers;
        public SearchForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.comLayerName.SelectedItem == null) {return;}
            if (this.ListFieldsName.SelectedItem == null) { return; }
            if (this.listBox1.SelectedItem == null) { return; }
            IFeatureLayer pFeatureLayer = GetLayerbyName(this.comLayerName.SelectedItem.ToString()) as IFeatureLayer;
            string Fname = (pFeatureLayer as ITable).Fields.get_Field((pFeatureLayer as ITable).Fields.FindFieldByAliasName(this.ListFieldsName.SelectedItem.ToString())).Name;
            string ret_sql = "";
            string ret_sql2 = "";
            if (Fname != "ZBGHSYSJ")
            { ret_sql = Fname + " LIKE '" + this.listBox1.SelectedItem.ToString() + "'";
            ret_sql2 = Fname + " LIKE '" + this.listBox1.SelectedItem.ToString() + "'";
            }
            else { ret_sql = Fname + " LIKE #" + this.listBox1.SelectedItem.ToString() + "#";
            ret_sql2 = Fname + " LIKE '" + this.listBox1.SelectedItem.ToString() + " 0:00:00'"; 
            }
            if (this.SqlOK != null)
                this.SqlOK(this, new SQLFileterEventArgs()
                {
                    SQL = ret_sql,
                    LayerIndex = this.comLayerName.SelectedIndex,
                    SQL_2 = ret_sql2
                });
        }

        public void ShowInfo(List<ILayer> layers)
        {
            if (layers.Count == 0) { return; }
            _layers = layers;
            this.comLayerName.Items.Clear();
            foreach (ILayer _layerItem in _layers)
            {
                this.comLayerName.Items.Add(_layerItem.Name);
            }
            comLayerName.SelectedItem = comLayerName.Items[0];
            base.Show();
        }

        private void comLayerName_SelectedIndexChanged(object sender, EventArgs e)
        {
            ITable _Table = GetLayerbyName (this.comLayerName.SelectedItem as string) as ITable;
            this.ListFieldsName.Items.Clear();
            for (int fieldIndex = 0;
                fieldIndex < _Table.Fields.FieldCount;
                fieldIndex++)
            {
                this.ListFieldsName.Items.Add
                     (_Table.Fields.Field[fieldIndex].AliasName);
            }
            ListFieldsName.SelectedItem = ListFieldsName.Items[0];
        }

        private ILayer GetLayerbyName(string layerName)
        {
            foreach (ILayer LayerItem in this._layers)
            {
                if (LayerItem.Name.Equals(layerName)) return LayerItem;
            }
            return null;
        }

        private void ListFieldsName_SelectedIndexChanged(object sender, EventArgs e)
        {
            ArrayList AL = GetLayerUniqueFieldValueByDataStatistics(GetLayerbyName(comLayerName.SelectedItem.ToString()) as IFeatureLayer,ListFieldsName.SelectedItem.ToString());
            listBox1.Items.Clear();
            foreach (object obj in AL)
            {
                this.listBox1.Items.Add(obj);
            }
        }

        private ArrayList GetLayerUniqueFieldValueByDataStatistics(IFeatureLayer pFeatureLayer, string fieldName)
        {
            ArrayList arrValues = new ArrayList();
            IQueryFilter pQueryFilter = new QueryFilterClass();
            IFeatureCursor pFeatureCursor = null;
            string Fname = (pFeatureLayer as ITable).Fields.get_Field((pFeatureLayer as ITable).Fields.FindFieldByAliasName(fieldName)).Name;
            pQueryFilter.SubFields = Fname;
            pFeatureCursor = pFeatureLayer.FeatureClass.Search(pQueryFilter, true);

            IDataStatistics pDataStati = new DataStatisticsClass();
            pDataStati.Field = Fname;
            pDataStati.Cursor = (ICursor)pFeatureCursor;

            IEnumerator pEnumerator = pDataStati.UniqueValues;
            pEnumerator.Reset();
            while (pEnumerator.MoveNext())
            {
                object pObj = pEnumerator.Current;
                arrValues.Add(pObj.ToString());
            }

            arrValues.Sort();
            return arrValues;
        }

        private void button1_Click(object sender, MouseEventArgs e)
        {
            if (this.comLayerName.SelectedItem == null) { return; }
            if (this.ListFieldsName.SelectedItem == null) { return; }
            if (this.listBox1.SelectedItem == null) { return; }
            IFeatureLayer pFeatureLayer = GetLayerbyName(this.comLayerName.SelectedItem.ToString()) as IFeatureLayer;
            string Fname = (pFeatureLayer as ITable).Fields.get_Field((pFeatureLayer as ITable).Fields.FindFieldByAliasName(this.ListFieldsName.SelectedItem.ToString())).Name;
            string ret_sql = "";
            string ret_sql2 = "";
            if (Fname != "ZBGHSYSJ")
            {
                ret_sql = Fname + " LIKE '" + this.listBox1.SelectedItem.ToString() + "'";
                ret_sql2 = Fname + " LIKE '" + this.listBox1.SelectedItem.ToString() + "'";
            }
            else
            {
                ret_sql = Fname + " LIKE #" + this.listBox1.SelectedItem.ToString() + "#";
                ret_sql2 = Fname + " LIKE '" + this.listBox1.SelectedItem.ToString() + " 0:00:00'";
            }
            if (this.SqlOK != null)
                this.SqlOK(this, new SQLFileterEventArgs()
                {
                    SQL = ret_sql,
                    LayerIndex = this.comLayerName.SelectedIndex,
                    SQL_2 = ret_sql2
                });
        }
    }
}
