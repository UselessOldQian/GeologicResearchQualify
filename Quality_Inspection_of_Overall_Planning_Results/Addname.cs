using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ESRI.ArcGIS.Controls;
using ESRI.ArcGIS.SystemUI;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.Geodatabase;
using QI_ClassLibrary;

namespace Quality_Inspection_of_Overall_Planning_Results
{
    public partial class Addname : Form
    {
        IFeatureLayer pFLayer;
        ZenJian zenjian;
        List<string> FieldName;
        public Addname(IFeatureLayer pFLayer,ZenJian zenjian,List<string> FieldName)
        {
            InitializeComponent();
            this.pFLayer = pFLayer;
            this.zenjian = zenjian;
            this.FieldName = FieldName;
        }

        

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.textBox1.Text != "" && this.textBox1.Text != null)
            {
                Revise re = new Revise();
                re.addField(pFLayer, this.textBox1.Text);
            }
            zenjian.dgvTable.DataSource = pFLayer;
            LoadData LD = new LoadData();
            zenjian.pDT = LD.ShowTableInDataGridView_zenjian(pFLayer as ITable, zenjian.dgvTable, out FieldName);
            base.Close();
            
        }
    }
}
