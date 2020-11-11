using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ESRI.ArcGIS.Controls;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Geodatabase;

namespace Quality_Inspection_of_Overall_Planning_Results
{
    public partial class AddSHP : Form
    {
        ZenJian zj;
        public AddSHP(ZenJian zj)
        {
            InitializeComponent();
            this.zj = zj;
            showInList(zj.axMapControl1);
        }

        private void showInList(AxMapControl mapControl)
        {
            for (int i = 0; i < mapControl.LayerCount; i++)
            {
                lb_main.Items.Add(mapControl.get_Layer(i).Name);
                lb_new.Items.Add(mapControl.get_Layer(i).Name);
            }
        }

        private void btn_ok_Click(object sender, EventArgs e)
        {
            if (lb_main.SelectedItem == null || lb_new.SelectedItem == null)
            {
                MessageBox.Show("请选择一个主图层和一个需要合并图层");
                return;
            }
            if (lb_main.SelectedItem.ToString() == lb_new.SelectedItem.ToString())
            {
                MessageBox.Show("请选择两个不同的图层");
                return;
            }
            IFeatureLayer fl_target = zj.GetLayerByName(lb_main.SelectedItem.ToString()) as IFeatureLayer;
            ISpatialReference pSourceSpr_target = ((fl_target.FeatureClass as IDataset) as IGeoDataset).SpatialReference;

            IFeatureLayer fl_new = zj.GetLayerByName(lb_new.SelectedItem.ToString()) as IFeatureLayer;
            ISpatialReference pSourceSpr_new = ((fl_new.FeatureClass as IDataset) as IGeoDataset).SpatialReference;
            if (fl_new == fl_target)
            {
                MessageBox.Show("投影坐标系不同，请转换成相同的");
                return;
            }
            IFeatureClass fc_new = fl_new.FeatureClass;
            int iFeaCnt = fc_new.FeatureCount(null);
            IFeatureClass fc_target = fl_target.FeatureClass;


            //获得源文件的游标
            IFeatureCursor pFeaCur = fc_new.Search(null, false);
            IFeature pSourceFea = pFeaCur.NextFeature();
            if (pSourceFea == null)
                return;




            bool blOK = false;
            string sError = "";

            try
            {
                //遍历每一条记录
                while (pSourceFea != null)
                {
                    //创建对应的目标记录
                    IFeature pTargetFea = fc_target.CreateFeature();
                    if (pTargetFea != null)
                    {
                        //复制几何
                        pTargetFea.Shape = pSourceFea.Shape;
                        //复制属性
                        CopyAttribute(pSourceFea, pTargetFea);
                        //保存编辑
                        pTargetFea.Store();
                    }
                    pSourceFea = pFeaCur.NextFeature();
                }
                blOK = true;
            }
            catch (Exception ex)
            {
                sError = ex.Message;
            }
            finally
            {
                if (blOK)
                {
                    MessageBox.Show("导入成功，共计导入：" + iFeaCnt + "条记录。");
                    this.Close();
                }
                else
                    MessageBox.Show("导入失败：" + sError);
            }
        }

        private void CopyAttribute(IFeature pSourceFea, IFeature pTargetFea)
        {
            //获取源文件的字段信息
            IFields pSourceFlds = pSourceFea.Fields;
            for (int i = 0; i < pSourceFlds.FieldCount; i++)
            {
                //筛掉OID 几何 大字段类型的字段名
                IField pSourceFld = pSourceFlds.get_Field(i);
                if ((pSourceFld.Type == esriFieldType.esriFieldTypeOID) ||
                    (pSourceFld.Type == esriFieldType.esriFieldTypeGeometry) ||
                    (pSourceFld.Type == esriFieldType.esriFieldTypeBlob))
                    continue;


                //筛掉值为空的
                string sSourceVal = pSourceFea.get_Value(i).ToString();
                if (string.IsNullOrEmpty(sSourceVal))
                    continue;


                //添加索引
                string sSourceFldName = pSourceFld.Name;
                int iTargetFldIndex = pTargetFea.Fields.FindField(sSourceFldName);
                if (iTargetFldIndex >= 0)
                    pTargetFea.set_Value(iTargetFldIndex, sSourceVal);
            }
        }  
    }
}
