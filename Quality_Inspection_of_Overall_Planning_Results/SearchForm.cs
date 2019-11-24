﻿using System;
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
            //if (this.listBox1.SelectedItem == null) { return; }
            IFeatureLayer pFeatureLayer = GetLayerbyName(this.comLayerName.SelectedItem.ToString()) as IFeatureLayer;
            string Fname = (pFeatureLayer as ITable).Fields.get_Field((pFeatureLayer as ITable).Fields.FindFieldByAliasName(this.ListFieldsName.SelectedItem.ToString())).Name;
            string ret_sql = "";
            string ret_sql2 = "";
            if (Fname != "ZBGHSYSJ")
            { 
                //ret_sql = Fname + " LIKE '" + this.listBox1.SelectedItem.ToString() + "'";
                ret_sql = Fname + " LIKE '*" + this.textBox1.Text + "*'";
                //ret_sql2 = Fname + " LIKE '" + this.listBox1.SelectedItem.ToString() + "'";
                ret_sql2 = Fname + " LIKE '*" + this.textBox1.Text + "*'";
            }
            else {
                ret_sql = Fname + " LIKE #" + this.listBox1.SelectedItem.ToString() + "#";
                ret_sql2 = Fname + " LIKE '" + this.listBox1.SelectedItem.ToString() + " 0:00:00'"; 
            }

            //测试取值
            int len = 0;
            //string combotext = this.panel1.Controls.Find("ComboBox1", false)[len].Text;


            foreach (Control c in this.panel1.Controls)
            {

                if (c.Name == "text1")
                {
                    string combotext = this.panel1.Controls.Find("ComboBox1", false)[len].Text;
                    string Fcombotext = (pFeatureLayer as ITable).Fields.get_Field((pFeatureLayer as ITable).Fields.FindFieldByAliasName(combotext)).Name;
                    len += 1;
                    string text = ((TextBox)c).Text;
                    ret_sql += " And " + Fcombotext + " LIKE '*" + text + "*'";
                    ret_sql2 += " And " + Fcombotext + " LIKE '*" + text + "*'";
                }

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
            ITable _Table = GetLayerbyName(this.comLayerName.SelectedItem as string) as ITable;
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
        //选中属性的变化
        private void ListFieldsName_SelectedIndexChanged(object sender, EventArgs e)
        {
            ArrayList AL = GetLayerUniqueFieldValueByDataStatistics(GetLayerbyName(comLayerName.SelectedItem.ToString()) as IFeatureLayer, ListFieldsName.SelectedItem.ToString());
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

        private void label3_Click(object sender, EventArgs e)
        {

        }
        
        int y1 = 54;
        int x1 = 18;
        int x2 = 65;
        int y2 = 51;
        int x3 = 330;
        int y3 = 57;
        int x4 = 365;
        int y4 = 50;
        int x5 = 585;
        int y5 = 49;
        private void button1_Click_1(object sender, EventArgs e)
        {
            //int x1 = this.label1.Location.X;
            //int y1 = this.label1.Location.Y;
            Label Lab1;
            ComboBox Combo;
            Label Lab2;
            TextBox text;
            Button But;
            //属性label
            x1 = 18;
            y1 += 30;
            Lab1 = new Label();
            Lab1.Size = new Size(41, 12);
            Lab1.Location = new System.Drawing.Point(x1, y1);
            Lab1.Text = "属性：";
            x1 += 80;
            panel1.Controls.Add(Lab1);
            //
            x2 = 65;
            y2 += 30;
            Combo = new ComboBox();
            Combo.Size = new Size(249, 20);
            Combo.Name = "ComboBox1";
            Combo.Location = new System.Drawing.Point(x2, y2);
            
            x2 += 80;
            panel1.Controls.Add(Combo);


            for (int fieldIndex = 0;
                        fieldIndex < ListFieldsName.Items.Count;
                        fieldIndex++)
            {
                Combo.Items.Add
                     (ListFieldsName.Items[fieldIndex]);
            }
            //ListFieldsName.SelectedItem = ListFieldsName.Items[0];
            //ITable _Table = GetLayerbyName(this.comLayerName.SelectedItem as string) as ITable;
            //this.ListFieldsName.Items.Clear();
            //for (int fieldIndex = 0;
            //    fieldIndex < _Table.Fields.FieldCount;
            //    fieldIndex++)
            //{
            //    this.ListFieldsName.Items.Add
            //         (_Table.Fields.Field[fieldIndex].AliasName);
            //}
            //ListFieldsName.SelectedItem = ListFieldsName.Items[0];
            //
            x3 = 330;
            y3 += 30;
            Lab2 = new Label();
            Lab2.Size = new Size(29, 12);
            Lab2.Location = new System.Drawing.Point(x3, y3);
            Lab2.Text = "like";
            x3 += 80;
            panel1.Controls.Add(Lab2);
            //
            x4 = 365;
            y4 += 30;
            text = new TextBox();
            text.Size = new Size(204, 21);
            text.Location = new System.Drawing.Point(x4, y4);
            text.Name = "text1";
            x3 += 80;
            panel1.Controls.Add(text);
            //


        }
    }
}
