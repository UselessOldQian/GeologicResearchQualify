using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Controls;
using ESRI.ArcGIS.SystemUI;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.DataSourcesFile;
using ESRI.ArcGIS.DataSourcesGDB;
using System.Windows.Forms;

namespace Quality_Inspection_of_Overall_Planning_Results
{
    class Revise
    {
         //定义新字段
        public void addField(IFeatureLayer pFeatureLayer, string field_name)
        {
            if (pFeatureLayer.FeatureClass.FindField(field_name) >= 0) { return; }
            IField pField = new FieldClass();

            //字段编辑
            IFieldEdit pFieldEdit = pField as IFieldEdit;

            //新建字段名
            pFieldEdit.Name_2 = field_name;

            //获取图层
            IFeatureClass pFeatureClass = pFeatureLayer.FeatureClass;
            IClass pTable = pFeatureClass as IClass;      //use ITable or IClass
            pTable.AddField(pFieldEdit);
            //set values of every feature's field-"name_cit" in the first layer   
            for (int i = 1; i <= pFeatureClass.FeatureCount(null); i++)
            {
                IFeature pFeature = pFeatureClass.GetFeature(i);
                pFeature.set_Value(pFeature.Fields.FindField(field_name), null);   //每个要素的“A”字段存储的都是“B”。
                pFeature.Store();
            }
        }

        //修改字段属性（重）
        public void ChangeFieldValue(IFeatureLayer mlayer, string pGetFieldName, string pGetFieldAliasName, int pGetFieldLength, esriFieldType FieldType, int FieldIndex,ZenJian pForm)
        {
            try
            {
                IFeatureLayer pFeatureLayer = mlayer as IFeatureLayer;
                if (pGetFieldName != null || pGetFieldAliasName != null && pGetFieldLength != 0)
                {
                    ITable pTable = pFeatureLayer.FeatureClass as ITable;
                    IField pField = new FieldClass();

                    IFieldEdit pFieldEdit = pField as IFieldEdit;//添加Temp字段
                    pFieldEdit.Name_2 = "temp";
                    pFieldEdit.AliasName_2 = pGetFieldAliasName;
                    pFieldEdit.Length_2 = pGetFieldLength;
                    pFieldEdit.Type_2 = FieldType;
                    pTable.AddField(pField);

                    VolFieldValue(FieldIndex, pFeatureLayer);//为Temp字段传入修改字段值
                    ISchemaLock pSchemaLock = pTable as ISchemaLock;
                    pSchemaLock.ChangeSchemaLock(esriSchemaLock.esriExclusiveSchemaLock);
                    pTable.DeleteField(pFeatureLayer.FeatureClass.Fields.get_Field(FieldIndex));//删除原修改字段
                    pSchemaLock.ChangeSchemaLock(esriSchemaLock.esriSharedSchemaLock);
                    pFieldEdit.Name_2 = pGetFieldName;
                    pFieldEdit.AliasName_2 = pGetFieldAliasName;
                    pFieldEdit.Length_2 = pGetFieldLength;
                    pFieldEdit.Type_2 = FieldType;
                    pTable.AddField(pField);//添加原字段同名字段

                    int tempIndex = 0;//获取创建Temp字段位置
                    for (int i = 1; i <= pFeatureLayer.FeatureClass.Fields.FieldCount; i++)
                    {
                        if (pFeatureLayer.FeatureClass.Fields.get_Field(i).Name == "temp")
                        {
                            tempIndex = i;
                        }
                    }

                    VolFieldValue(tempIndex,pFeatureLayer);//为创建修改字段同名字段赋值
                    pSchemaLock.ChangeSchemaLock(esriSchemaLock.esriExclusiveSchemaLock);
                    pTable.DeleteField(pFeatureLayer.FeatureClass.Fields.get_Field(tempIndex));//删除创建Temp 字段
                    pSchemaLock.ChangeSchemaLock(esriSchemaLock.esriSharedSchemaLock);
                }
                else
                    return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "提示信息");
            }
        }

        // 字段集体赋值函数
        private void VolFieldValue(int FiledIndex, IFeatureLayer pFeatureLayer)
        {
            try
            {
                IFeatureCursor pFeatureCursor;
                pFeatureCursor = pFeatureLayer.FeatureClass.Search(null, false);
                IFeature pFeature;
                pFeature = pFeatureCursor.NextFeature();
                int pFieldCount = pFeatureLayer.FeatureClass.Fields.FieldCount;
                while (pFeature != null)
                {
                    if (pFeatureLayer.FeatureClass.Fields.get_Field(FiledIndex).ToString() == "Shape")
                    {
                        pFeature.set_Value(pFieldCount - 1, pFeature.Shape.GeometryType.ToString());
                    }
                    else
                        pFeature.set_Value(pFieldCount - 1, pFeature.get_Value(FiledIndex));
                    pFeature.Store();
                    pFeature = pFeatureCursor.NextFeature();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "提示信息");
            }
        }

        // 删除字段函数
        public void deleteField(IFeatureLayer pFeatureLayer, string field_name)
        {
            try
            {
                ITable pDelTable = pFeatureLayer.FeatureClass as ITable;
                ISchemaLock pSchemaLock = (ISchemaLock)pDelTable;
                pSchemaLock.ChangeSchemaLock(esriSchemaLock.esriExclusiveSchemaLock);
                IField pDelField = pDelTable.Fields.get_Field(pDelTable.Fields.FindField(field_name));//实现数据独占，避免数据使用冲突，只对Geodatabase有效
                pDelTable.DeleteField(pDelField);
                pSchemaLock.ChangeSchemaLock(esriSchemaLock.esriSharedSchemaLock);//释放占有资源
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "提示信息");
            }
        }

        /// <summary>
        /// 删除字段值
        /// </summary>
        /// <param name="pCurrentLayer"></param>
        /// <param name="fieldName"></param>
        public static bool DeleteILayerField(IFeatureLayer pCurrentLayer, string fieldName)
        {
            try
            {
                IFeatureLayer pFeatureLayer = pCurrentLayer;
                IFeatureClass pFeatureClass = pFeatureLayer.FeatureClass;
                //
                int pFieldIndex = pFeatureClass.FindField(fieldName);
                IFields pFields = pFeatureClass.Fields;
                IField pField = pFields.get_Field(pFieldIndex);
                if (pField != null)
                {
                    pFeatureClass.DeleteField(pField);
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
    }
}
