using System;
using System.Collections.Generic;
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
using System.Data.OleDb;
using System.Data.SqlClient;


namespace QI_ClassLibrary
{
    public class CheckDataConsistent
    {
        public string GetChineseName(string EnglishName)
        {
            switch (EnglishName)
            {
                case "XZQ":
                    return EnglishName + "(行政区)";
                case "JQDLTB":
                    return EnglishName + "(基期地类图斑)";
                case "CSKFBJNGHYT":
                    return EnglishName + "(城市开发边界内规划用途)";
                case "CSKFBJ":
                    return EnglishName + "(城市开发边界)";
                case "JSYDKZX":
                    return EnglishName + "(建设用地控制线)";
                case "STKJKZX":
                    return EnglishName + "(生态空间控制线)";
                case "JSYDHJBNTGZ2035":
                    return EnglishName + "(建设用地和基本农田管制2035)";
                case "JLHDK":
                    return EnglishName + "(简化量地块)";
                case "TDLYJGTZB":
                    return EnglishName + "(土地利用结构调整表)";
                case "GDZBPHB":
                    return EnglishName + "(耕地占补平衡表)";
                default:
                    return EnglishName;
            }
        }


        /// <summary>
        /// 检查输入的图层组空间数据范围是否一致
        /// </summary>
        /// <param name="Layers"></param>
        /// <param name="pErrorDataTable"></param>
        /// <returns></returns>
        public string CheckSpatialRangeConsistent1(ILayer[] Layers, ref DataTable pErrorDataTable)
        {
            string ERROR = "";
            try
            {
                if (Layers.Length < 2) { return null; }
                for (int i = 0; i < Layers.Length; i++)
                {
                    if (Layers[i] == null) { return ERROR; }
                    IEnvelope pEnv1 = ((Layers[i] as FeatureLayer).FeatureClass as IGeoDataset).Extent;
                    ISpatialReference GRout = ((Layers[i] as IFeatureLayer).FeatureClass as IGeoDataset).SpatialReference;
                    IGeometry pGeoOut = UnionAll((Layers[i] as IFeatureLayer).FeatureClass) as IGeometry;
                    for (int j = i + 1; j < Layers.Length; j++)
                    {
                        ISpatialReference GRin = ((Layers[j] as IFeatureLayer).FeatureClass as IGeoDataset).SpatialReference;
                        if (GRout.SpatialReferenceImpl != GRin.SpatialReferenceImpl || GRout.Name != GRin.Name)  
                        {
                            MessageBox.Show("外部数据" + Layers[j].Name + "与" + Layers[i].Name + "的坐标系不同，" + Layers[j].Name + "与" + Layers[i].Name + "范围无法比较，请修改"); 
                        }
                        IGeometry pGeoIn = UnionAll((Layers[j] as IFeatureLayer).FeatureClass) as IGeometry;
                        IEnvelope pEnv2 = (Layers[j] as IGeoDataset).Extent;
                        ITopologicalOperator pGeoOutTP = pGeoOut as ITopologicalOperator;
                        IGeometry pDiff = pGeoOutTP.Difference(pGeoIn);
                        IArea pArea = pDiff as IArea;
                        if (pArea.Area!=0)
                        {
                            ERROR += "\r\nERROR6401:" + GetChineseName(Layers[i].Name) + "与" + GetChineseName(Layers[j].Name) + "数据范围不一致," + GetChineseName(Layers[i].Name) + "减去与" + GetChineseName(Layers[j].Name) + "相交面积后，相差面积为" + pArea.Area + "平方米";
                            pErrorDataTable.Rows.Add(new object[] { "6401", Layers[i].Name, null, null, GetChineseName(Layers[i].Name) + "与" + GetChineseName(Layers[j].Name) + "数据范围不一致,相差面积为" + pArea.Area + "平方米", false, true });
                        }
                    }
                }
                return ERROR;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return ERROR;
            }
        }

        /// <summary>
        /// 判断InData是否在OutData内
        /// </summary>
        /// <param name="OutData"></param>
        /// <param name="InData"></param>
        /// <param name="pErrorDataTable"></param>
        /// <returns></returns>
        public string CheckSpatialRangeConsistent2(IFeatureLayer OutData, IFeatureLayer InData, ref DataTable pErrorDataTable, string Layername, string ErrorNumber)
        {
            if (InData == null) { return ""; }
            if (OutData == null) { MessageBox.Show("外部数据" + Layername + "不存在,请加入相应数据"); return "\r\n外部数据" + Layername + "不存在"; }
            ISpatialReference GRout = (OutData.FeatureClass as IGeoDataset).SpatialReference;
            ISpatialReference GRin = (InData.FeatureClass as IGeoDataset).SpatialReference;
            if (GRout.SpatialReferenceImpl != GRin.SpatialReferenceImpl || GRout.Name != GRin.Name)
            { MessageBox.Show("外部数据" + OutData.Name + "与" + InData.Name + "的坐标系不同，" + InData.Name + "与" + OutData.Name + "范围无法比较，请修改"); return "\r\n外部数据" + OutData.Name + "与" + InData.Name + "的坐标系不同"; }
            //IRelationalOperator pGeoOut = UnionAll(OutData.FeatureClass) as IRelationalOperator;
            //IGeometry pGeoIn = UnionAll(InData.FeatureClass);
            IGeometry pGeoOut = UnionAll(OutData.FeatureClass) as IGeometry;
            //IGeometry pGeoOut = UnionAll(OutData.FeatureClass) as IGeometry;
            IGeometry pGeoIn = UnionAll(InData.FeatureClass) as IGeometry;
            ITopologicalOperator pGeoInTP = pGeoIn as ITopologicalOperator;
            IGeometry pDiff = pGeoInTP.Difference(pGeoOut);
            IArea pArea = pDiff as IArea;
            if (pArea.Area != 0)//pGeoOut.Contains(pGeoIn))
            {
                pErrorDataTable.Rows.Add(new object[] { ErrorNumber, InData.Name, null, null, GetChineseName(InData.Name) + "不在外部数据" + GetChineseName(OutData.Name) + "范围内,不在范围内的面积为" + pArea.Area + "平方米", false, true });
                return "\r\nERROR" + ErrorNumber + ":" + GetChineseName(InData.Name) + "不在外部数据" + GetChineseName(OutData.Name) + "范围内,不在范围内的面积为" + pArea.Area + "平方米";
            }
            return null;
        }

        /// <summary>
        /// 判断Data1和Data2的Geometry是否包含
        /// </summary>
        /// <param name="Data1">外部数据</param>
        /// <param name="Data2"></param>
        /// <param name="pErrorDataTable"></param>
        /// <returns></returns>
        public string CheckSpatialRangeEquals(IFeatureLayer Data1, IFeatureLayer Data2, ref DataTable pErrorDataTable, string Layername, string ErrorNumber)
        {
            if (Data2 == null) { return ""; }
            if (Data1 == null) { MessageBox.Show("外部数据" + Layername + "不存在,请加入相应数据"); return "\r\n外部数据" + Layername + "不存在"; }
            ISpatialReference GRout = (Data1.FeatureClass as IGeoDataset).SpatialReference;
            ISpatialReference GRin = (Data2.FeatureClass as IGeoDataset).SpatialReference;
            if (GRout.SpatialReferenceImpl != GRin.SpatialReferenceImpl || GRout.Name != GRin.Name)
            { MessageBox.Show("外部数据" + Data1.Name + "与" + Data2.Name + "的坐标系不同，" + Data2.Name + "与" + Data1.Name + "范围无法比较，请修改"); return "\r\n外部数据" + Data1.Name + "与" + Data2.Name + "的坐标系不同"; }
            IGeometry Geo2 = UnionAll(Data2.FeatureClass);
            //string relationDescription = "RELATE(G1, G2, '**T******')";
            ISpatialFilter pSpatialFilter = new SpatialFilterClass();
            pSpatialFilter.Geometry = Geo2 as IGeometry;
            pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects;
            IGeometry Geo1 = UnionAllSelect(Data1.FeatureClass, pSpatialFilter);

            ITopologicalOperator pGeoInTP = Geo2 as ITopologicalOperator;
            IGeometry pDiff = pGeoInTP.Difference(Geo1);
            IArea pArea = pDiff as IArea;
            //IRelationalOperator RO = Geo2 as IRelationalOperator;
            //bool isEqual = RO.Within(Geo1);//.Relation(Geo1, relationDescription);
            if (Geo1 != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Geo1);
            if (pGeoInTP != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pGeoInTP);
            System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString() + " end diff");
            if (pArea.Area != 0)
            {
                pErrorDataTable.Rows.Add(new object[] { ErrorNumber, Data2.Name, null, null, GetChineseName(Data2.Name) + "不在外部数据" + GetChineseName(Data1.Name) + "范围内,不在范围内的面积为" + pArea.Area + "平方米", false, true });
                return "\r\nERROR" + ErrorNumber + ":" + GetChineseName(Data2.Name) + "不在外部数据" + GetChineseName(Data1.Name) + "范围内,不在范围内的面积为" + pArea.Area + "平方米";
            }
            return null;
        }

        /// <summary>
        /// condition为Data2的属性查询条件
        /// </summary>
        /// <param name="Data1"></param>
        /// <param name="Data2"></param>
        /// <param name="pErrorDataTable"></param>
        /// <param name="Layername"></param>
        /// <param name="condition"></param>
        /// <returns></returns>
        public string CheckSpatialRangeEquals(IFeatureLayer Data1, IFeatureLayer Data2, ref DataTable pErrorDataTable, string Layername, string condition, string ErrorNumber)
        {
            if (Data2 == null) { return ""; }
            if (Data1 == null) { return ""; }
            ISpatialReference GRout = (Data1.FeatureClass as IGeoDataset).SpatialReference;
            ISpatialReference GRin = (Data2.FeatureClass as IGeoDataset).SpatialReference;
            if (GRout.SpatialReferenceImpl != GRin.SpatialReferenceImpl || GRout.Name != GRin.Name)
            { MessageBox.Show("外部数据" + Data1.Name + "与" + Data2.Name + "的坐标系不同，" + Data2.Name + "与" + Data1.Name + "范围无法比较，请修改"); return "\r\n" + Data1.Name + "与" + Data2.Name + "的坐标系不同"; }

            //string relationDescription = "RELATE(G1, G2, '**T******')";
            IQueryFilter pQueryFilter = new QueryFilterClass();
            pQueryFilter.WhereClause = condition;
            IGeometry Geo2 = UnionAllSelect(Data2.FeatureClass, pQueryFilter);
            ISpatialFilter pSpatialFilter = new SpatialFilterClass();
            pSpatialFilter.Geometry = Geo2 as IGeometry;
            pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects;
            IGeometry Geo1 = UnionAllSelect(Data1.FeatureClass, pSpatialFilter);

            ITopologicalOperator pGeoInTP = Geo2 as ITopologicalOperator;
            IGeometry pDiff = pGeoInTP.Difference(Geo1);
            IArea pArea = pDiff as IArea;
            //IRelationalOperator RO = Geo2 as IRelationalOperator;
            //bool isEqual = RO.Within(Geo1);//.Relation(Geo1, relationDescription);
            if (Geo1 != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Geo1);
            if (pGeoInTP != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pGeoInTP);
            System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString() + " end diff");
            if (pArea.Area != 0)
            {
                pErrorDataTable.Rows.Add(new object[] { ErrorNumber, Data2.Name, null, null, GetChineseName(Data2.Name) + " "+ condition+" 没有全部位于" + GetChineseName(Data1.Name) + "范围内,不在范围内的面积为" + pArea.Area + "平方米", false, true });
                return "\r\nERROR" + ErrorNumber + ":" + GetChineseName(Data2.Name) + " " + condition + " 没有全部位于" + GetChineseName(Data1.Name) + "范围内,不在范围内的面积为" + pArea.Area + "平方米";
            }
            return null;
        }


        public string CheckSpatialRangeEquals(IFeatureLayer Data1, IFeatureLayer Data2, ref DataTable pErrorDataTable, string Layername, string condition, string Layername2, string condition2, string ErrorNumber)
        {
            if (Data2 == null) { return ""; }
            if (Data1 == null && Layername != null) { MessageBox.Show("外部数据" + Layername + "不存在,请加入相应数据"); return "\r\n外部数据" + Layername + "不存在"; }
            ISpatialReference GRout = (Data1.FeatureClass as IGeoDataset).SpatialReference;
            ISpatialReference GRin = (Data2.FeatureClass as IGeoDataset).SpatialReference;
            if (GRout.SpatialReferenceImpl != GRin.SpatialReferenceImpl || GRout.Name != GRin.Name)
            { MessageBox.Show("外部数据" + Data1.Name + "与" + Data2.Name + "的坐标系不同，" + Data2.Name + "与" + Data1.Name + "范围无法比较，请修改"); return "\r\n外部数据" + Data1.Name + "与" + Data2.Name + "的坐标系不同"; }

            //string relationDescription = "RELATE(G1, G2, '**T******')";
            IQueryFilter pQueryFilter = new QueryFilterClass();
            pQueryFilter.WhereClause = condition;
            IGeometry Geo1 = UnionAllSelect(Data1.FeatureClass, pQueryFilter);
            IQueryFilter pQueryFilter2 = new QueryFilterClass();
            pQueryFilter2.WhereClause = condition2;
            IGeometry Geo2 = UnionAllSelect(Data2.FeatureClass, pQueryFilter2);
            ISpatialFilter pSpatialFilter = new SpatialFilterClass();

            ITopologicalOperator pGeoInTP = Geo2 as ITopologicalOperator;
            IGeometry pDiff = pGeoInTP.Difference(Geo1);
            IArea pArea = pDiff as IArea;
            //IRelationalOperator RO = Geo2 as IRelationalOperator;
            //bool isEqual = RO.Within(Geo1);//.Relation(Geo1, relationDescription);
            if (Geo1 != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Geo1);
            if (pGeoInTP != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pGeoInTP);
            System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString() + " end diff");
            if (pArea.Area != 0)
            {
                pErrorDataTable.Rows.Add(new object[] { ErrorNumber, Data2.Name, null, null, GetChineseName(Data1.Name) + " " + condition + " 与" + GetChineseName(Data2.Name) + " " + condition2 + " 范围不一致," + GetChineseName(Data1.Name) + "图层 不在" + GetChineseName(Data2.Name) + "图层 范围内的面积为" + pArea.Area + "平方米", false, true });
                return "\r\nERROR" + ErrorNumber + ":" + GetChineseName(Data1.Name) + " " + condition + " 与" + GetChineseName(Data2.Name) + " " + condition2 + " 范围不一致," + GetChineseName(Data1.Name) + "图层 不在" + GetChineseName(Data2.Name) + "图层 范围内的面积为" + pArea.Area + "平方米";
            }
            return null;
        }

        public string CheckSpatialRangeEquals_2035_6501(IFeatureLayer Data1, IFeatureLayer Data2, ref DataTable pErrorDataTable, string Layername, string condition, string Layername2, string condition2, string ErrorNumber)
        {
            if (Data2 == null) { return ""; }
            if (Data1 == null && Layername != null) { MessageBox.Show("外部数据" + Layername + "不存在,请加入相应数据"); return "\r\n外部数据" + Layername + "不存在"; }
            ISpatialReference GRout = (Data1.FeatureClass as IGeoDataset).SpatialReference;
            ISpatialReference GRin = (Data2.FeatureClass as IGeoDataset).SpatialReference;
            if (GRout.SpatialReferenceImpl != GRin.SpatialReferenceImpl || GRout.Name != GRin.Name)
            { MessageBox.Show("外部数据" + Data1.Name + "与" + Data2.Name + "的坐标系不同，" + Data2.Name + "与" + Data1.Name + "范围无法比较，请修改"); return "\r\n外部数据" + Data1.Name + "与" + Data2.Name + "的坐标系不同"; }

            //string relationDescription = "RELATE(G1, G2, '**T******')";
            IQueryFilter pQueryFilter = new QueryFilterClass();
            pQueryFilter.WhereClause = condition;
            IGeometry Geo1 = UnionAllSelect(Data1.FeatureClass, pQueryFilter);
            IQueryFilter pQueryFilter2 = new QueryFilterClass();
            pQueryFilter2.WhereClause = condition2;
            IGeometry Geo2 = UnionAllSelect(Data2.FeatureClass, pQueryFilter2);
            ISpatialFilter pSpatialFilter = new SpatialFilterClass();

            ITopologicalOperator pGeoInTP = Geo1 as ITopologicalOperator;
            IGeometry pDiff = pGeoInTP.Difference(Geo2);
            IArea pArea = pDiff as IArea;
            //IRelationalOperator RO = Geo2 as IRelationalOperator;
            //bool isEqual = RO.Within(Geo1);//.Relation(Geo1, relationDescription);
            if (Geo2 != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Geo1);
            if (pGeoInTP != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pGeoInTP);
            System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString() + " end diff");
            if (pArea.Area != 0)
            {
                pErrorDataTable.Rows.Add(new object[] { ErrorNumber, Data2.Name, null, null, GetChineseName(Data1.Name) + " " + condition + " 与" + GetChineseName(Data2.Name) + " " + condition2 + " 范围不一致," + GetChineseName(Data1.Name) + "图层 不在" + GetChineseName(Data2.Name) + "图层 范围内的面积为" + pArea.Area + "平方米", false, true });
                return "\r\nERROR" + ErrorNumber + ":" + GetChineseName(Data1.Name) + " " + condition + " 与" + GetChineseName(Data2.Name) + " " + condition2 + " 范围不一致," + GetChineseName(Data1.Name) + "图层 不在" + GetChineseName(Data2.Name) + "图层 范围内的面积为" + pArea.Area + "平方米";
            }
            return null;
        }

        public string CheckSpatialRangeEquals(IFeatureLayer Data1, IFeatureLayer Data2, ref DataTable pErrorDataTable, string Layername, string condition, string Layername2, string condition2, string relationDescription, string ErrorNumber)
        {
            if (Data2 == null) { return ""; }
            if (Data1 == null && Layername != null) { MessageBox.Show(Layername + "不存在,请加入相应数据"); return "\r\n" + Layername + "不存在"; }
            ISpatialReference GRout = (Data1.FeatureClass as IGeoDataset).SpatialReference;
            ISpatialReference GRin = (Data2.FeatureClass as IGeoDataset).SpatialReference;
            if (GRout.SpatialReferenceImpl != GRin.SpatialReferenceImpl || GRout.Name != GRin.Name)
            { MessageBox.Show("外部数据" + Data1.Name + "与" + Data2.Name + "的坐标系不同，" + Data2.Name + "与" + Data1.Name + "范围无法比较，请修改"); return "\r\n外部数据" + Data1.Name + "与" + Data2.Name + "的坐标系不同"; }

            //string relationDescription = "RELATE(G1, G2, '**T******')";
            IQueryFilter pQueryFilter = new QueryFilterClass();
            pQueryFilter.WhereClause = condition;
            IGeometry Geo1 = UnionAllSelect(Data1.FeatureClass, pQueryFilter);
            IQueryFilter pQueryFilter2 = new QueryFilterClass();
            pQueryFilter.WhereClause = condition2;
            IGeometry Geo2 = UnionAllSelect(Data2.FeatureClass, pQueryFilter2);
            ISpatialFilter pSpatialFilter = new SpatialFilterClass();

            //ITopologicalOperator pGeoInTP = Geo2 as ITopologicalOperator;
            //IGeometry pDiff = pGeoInTP.Difference(Geo1);
            //IArea pArea = pDiff as IArea;
            IRelationalOperator RO = Geo1 as IRelationalOperator;
            bool isRelation = RO.Relation(Geo2, relationDescription);
            if (!isRelation)
            {
                if (relationDescription == "RELATE(G1, G2, 'T*F*T*F**')")
                {
                    ITopologicalOperator pTO = Geo1 as ITopologicalOperator;
                    IGeometry diff = pTO.Difference(Geo2);
                    if (Geo2 != null)
                        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Geo2);
                    if (RO != null)
                        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(RO);
                    if ((diff as IArea).Area != 0)
                    {
                        pErrorDataTable.Rows.Add(new object[] { ErrorNumber, Data1.Name, null, null, GetChineseName(Data1.Name) + " " + condition + " 与" + GetChineseName(Data2.Name) + " " + condition2 + " 范围不一致，" + GetChineseName(Data1.Name) + " 不在" + GetChineseName(Data1.Name) + " 内的面积为" + (diff as IArea).Area + "平方米", false, true });
                        return "\r\nERROR" + ErrorNumber + ":" + GetChineseName(Data1.Name) + " " + condition + " 与" + GetChineseName(Data2.Name) + " " + condition2 + " 范围不一致，" + GetChineseName(Data1.Name) + " 不在" + GetChineseName(Data1.Name) + " 内的面积为" + (diff as IArea).Area + "平方米";
                    }
                }
                else
                {
                    pErrorDataTable.Rows.Add(new object[] { ErrorNumber, Data1.Name, null, null, GetChineseName(Data1.Name) + " " + condition + " 未全部位于" + GetChineseName(Data2.Name)  + " " + condition2 + " 范围外", false, true });
                    return "\r\nERROR" + ErrorNumber + ":" + GetChineseName(Data1.Name) + " " + condition + " 未全部位于" + GetChineseName(Data2.Name) + " " + condition2 + " 范围外";
                }
            }
            if (Geo2 != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Geo2);
            if (RO != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(RO);
            return null;
        }

        public string CheckSpatialRangeNotWithin(IFeatureLayer Data1, IFeatureLayer Data2, ref DataTable pErrorDataTable,string condition)
        {
            if (Data2 == null) { return ""; }
            if (Data1 == null) { return ""; }
            ISpatialReference GRout = (Data1.FeatureClass as IGeoDataset).SpatialReference;
            ISpatialReference GRin = (Data2.FeatureClass as IGeoDataset).SpatialReference;
            if (GRout.SpatialReferenceImpl != GRin.SpatialReferenceImpl || GRout.Name != GRin.Name)
            { MessageBox.Show("外部数据" + Data1.Name + "与" + Data2.Name + "的坐标系不同，" + Data2.Name + "与" + Data1.Name + "范围无法比较，请修改"); return "\r\n外部数据" + Data1.Name + "与" + Data2.Name + "的坐标系不同"; }

            //string relationDescription = "RELATE(G1, G2, '**T******')";
            IQueryFilter pQueryFilter = new QueryFilterClass();
            pQueryFilter.WhereClause = condition;
            IGeometry Geo1 = UnionAllSelect(Data1.FeatureClass, pQueryFilter);
            pQueryFilter.WhereClause = "LX LIKE '城市开发边界内建设用地'";
            IGeometry Geo2 = UnionAllSelect(Data2.FeatureClass, pQueryFilter);
            pQueryFilter.WhereClause = "LX LIKE '其他建设用地区'";
            IGeometry Geo3 = UnionAllSelect(Data2.FeatureClass, pQueryFilter);
            ISpatialFilter pSpatialFilter = new SpatialFilterClass();

            ITopologicalOperator pGeoInTP = Geo1 as ITopologicalOperator;
            IGeometry pIntersect1 = pGeoInTP.Intersect(Geo2,esriGeometryDimension.esriGeometry2Dimension);
            IGeometry pIntersect2 = pGeoInTP.Intersect(Geo3, esriGeometryDimension.esriGeometry2Dimension);
            IArea pArea1 = pIntersect1 as IArea;
            IArea pArea2 = pIntersect2 as IArea;
            //IRelationalOperator RO = Geo2 as IRelationalOperator;
            //bool isEqual = RO.Within(Geo1);//.Relation(Geo1, relationDescription);
            if (Geo1 != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Geo1);
            if (Geo2 != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Geo2);
            if (pGeoInTP != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pGeoInTP);
            if (pArea1.Area != 0||pArea2.Area != 0)
            {
                string Error=GetChineseName(Data2.Name) + "与" + GetChineseName(Data1.Name) + "相交";
                if (pArea1.Area != 0)
                {
                    Error += ",JSYDKZX图层类型为'城市开发边界内建设用地'的面积为"+pArea1.Area+"平方米";
                }
                if (pArea2.Area != 0)
                {
                    Error += ",JSYDKZX图层类型为'其他建设用地区'的面积为"+pArea2.Area+"平方米";
                }
                pErrorDataTable.Rows.Add(new object[] { "6502", Data2.Name, null, null, Error, false, true });
                return "\r\nERROR6502:" + Error;
            }
            return null;
        }

        public string CheckSpatialRangeNotWithin6503(IFeatureLayer Data1, IFeatureLayer Data2, ref DataTable pErrorDataTable)
        {
            if (Data2 == null) { return ""; }
            if (Data1 == null) { return ""; }
            ISpatialReference GRout = (Data1.FeatureClass as IGeoDataset).SpatialReference;
            ISpatialReference GRin = (Data2.FeatureClass as IGeoDataset).SpatialReference;
            if (GRout.SpatialReferenceImpl != GRin.SpatialReferenceImpl || GRout.Name != GRin.Name)
            { MessageBox.Show("外部数据" + Data1.Name + "与" + Data2.Name + "的坐标系不同，" + Data2.Name + "与" + Data1.Name + "范围无法比较，请修改"); return "\r\n外部数据" + Data1.Name + "与" + Data2.Name + "的坐标系不同"; }
            IQueryFilter pQueryFilter = new QueryFilterClass();
            IGeometry Geo1 = UnionAllSelect(Data1.FeatureClass, pQueryFilter);
            pQueryFilter.WhereClause = "GKDJ LIKE '01'";
            IGeometry Geo2 = UnionAllSelect(Data2.FeatureClass, pQueryFilter);
            pQueryFilter.WhereClause = "GKDJ LIKE '02'";
            IGeometry Geo3 = UnionAllSelect(Data2.FeatureClass, pQueryFilter);
            pQueryFilter.WhereClause = "GKDJ LIKE '03'";
            IGeometry Geo4 = UnionAllSelect(Data2.FeatureClass, pQueryFilter);
            ISpatialFilter pSpatialFilter = new SpatialFilterClass();

            ITopologicalOperator pGeoInTP = Geo1 as ITopologicalOperator;
            IGeometry pIntersect1 = pGeoInTP.Intersect(Geo2, esriGeometryDimension.esriGeometry2Dimension);
            IGeometry pIntersect2 = pGeoInTP.Intersect(Geo3, esriGeometryDimension.esriGeometry2Dimension);
            IGeometry pIntersect3 = pGeoInTP.Intersect(Geo4, esriGeometryDimension.esriGeometry2Dimension);
            IArea pArea1 = pIntersect1 as IArea;
            IArea pArea2 = pIntersect2 as IArea;
            IArea pArea3 = pIntersect3 as IArea;
            //IRelationalOperator RO = Geo2 as IRelationalOperator;
            //bool isEqual = RO.Within(Geo1);//.Relation(Geo1, relationDescription);
            if (Geo1 != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Geo1);
            if (Geo2 != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Geo2);
            if (Geo3 != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Geo3);
            if (Geo4 != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Geo4);
            if (pGeoInTP != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pGeoInTP);
            if (pArea1.Area != 0 || pArea2.Area != 0)
            {
                string Error = GetChineseName(Data2.Name) + "与" + GetChineseName(Data1.Name) + "相交";
                if (pArea1.Area != 0)
                {
                    Error += ",STKJKZX图层类型为'01'的面积为" + pArea1.Area + "平方米";
                }
                if (pArea2.Area != 0)
                {
                    Error += ",STKJKZX图层类型为'02'的面积为" + pArea2.Area + "平方米";
                }
                if (pArea3.Area != 0)
                {
                    Error += ",STKJKZX图层类型为'03'的面积为" + pArea3.Area + "平方米";
                }
                pErrorDataTable.Rows.Add(new object[] { "6503", Data2.Name, null, null, Error, false, true });
                return "\r\nERROR6503:" + Error;
            }
            return null;
        }

        /// <summary>
        /// 由于速度太慢暂时不把Geometry进行ConstructUnion，而是直接输出pGeometryCollection但如果需要计算面积的话只能这样走
        /// </summary>
        /// <param name="featureClass"></param>
        /// <returns></returns>
        public IGeometry UnionAll(IFeatureClass featureClass)
        {
            System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString() + " begin UnionAll");
            IFeatureCursor pFCursor = featureClass.Search(null, false);
            IFeature pFeature = pFCursor.NextFeature();    // 第一个Feature
            IGeometryBag pGeometryBag = new GeometryBag() as IGeometryBag;
            pGeometryBag.SpatialReference = (featureClass as IGeoDataset).SpatialReference;
            IGeometryCollection pGeometryCollection = pGeometryBag as IGeometryCollection;
            List<IFeature> featureList = new List<IFeature>();
            object missing = Type.Missing;
            int featureNum = 0;
            while (pFeature != null) // 如果是null就说明刚才那个是最后一个了，后面没有了
            {
                featureNum++;
                featureList.Add(pFeature);
                pGeometryCollection.AddGeometry(pFeature.ShapeCopy, missing, missing);
                System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString() + "  " + pFeature.OID);
                pFeature = pFCursor.NextFeature();
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pFCursor);
            ITopologicalOperator pTopologicalOperator = new Polygon() as ITopologicalOperator;
            pTopologicalOperator.ConstructUnion(pGeometryCollection as IEnumGeometry);
            System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString() + "end");
            IGeometry pGeometry = pTopologicalOperator as IGeometry; //pGeometryCollection as IGeometry;
            return pGeometry;// pTopologicalOperator as IGeometry;
        }


        public IGeometry UnionAllSelect(IFeatureClass featureClass, ISpatialFilter pSpatialFilter)
        {
            System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString() + " begin UnionAll");
            IFeatureCursor pFCursor = featureClass.Search(pSpatialFilter, false);
            IFeature pFeature = pFCursor.NextFeature();    // 第一个Feature
            IGeometryBag pGeometryBag = new GeometryBag() as IGeometryBag;
            pGeometryBag.SpatialReference = (featureClass as IGeoDataset).SpatialReference;
            IGeometryCollection pGeometryCollection = pGeometryBag as IGeometryCollection;
            List<IFeature> featureList = new List<IFeature>();
            object missing = Type.Missing;
            int featureNum = 0;
            while (pFeature != null) // 如果是null就说明刚才那个是最后一个了，后面没有了
            {
                featureNum++;
                featureList.Add(pFeature);
                pGeometryCollection.AddGeometry(pFeature.ShapeCopy, missing, missing);
                System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString() + "  " + pFeature.OID);
                pFeature = pFCursor.NextFeature();
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pFCursor);
            ITopologicalOperator pTopologicalOperator = new Polygon() as ITopologicalOperator;
            pTopologicalOperator.ConstructUnion(pGeometryCollection as IEnumGeometry);
            System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString() + "end");
            IGeometry pGeometry = pTopologicalOperator as IGeometry; //pGeometryCollection as IGeometry;
            return pGeometry;// pTopologicalOperator as IGeometry;
        }

        public IGeometry UnionAllSelect(IFeatureClass featureClass, IQueryFilter pQueryFilter)
        {
            System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString() + " begin UnionAll");
            IFeatureCursor pFCursor = featureClass.Search(pQueryFilter, false);
            IFeature pFeature = pFCursor.NextFeature();    // 第一个Feature
            IGeometryBag pGeometryBag = new GeometryBag() as IGeometryBag;
            pGeometryBag.SpatialReference = (featureClass as IGeoDataset).SpatialReference;
            IGeometryCollection pGeometryCollection = pGeometryBag as IGeometryCollection;
            List<IFeature> featureList = new List<IFeature>();
            object missing = Type.Missing;
            int featureNum = 0;
            while (pFeature != null) // 如果是null就说明刚才那个是最后一个了，后面没有了
            {
                featureNum++;
                featureList.Add(pFeature);
                pGeometryCollection.AddGeometry(pFeature.ShapeCopy, missing, missing);
                System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString() + "  " + pFeature.OID);
                pFeature = pFCursor.NextFeature();
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pFCursor);
            ITopologicalOperator pTopologicalOperator = new Polygon() as ITopologicalOperator;
            pTopologicalOperator.ConstructUnion(pGeometryCollection as IEnumGeometry);
            System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString() + "end");
            IGeometry pGeometry = pTopologicalOperator as IGeometry; //pGeometryCollection as IGeometry;
            return pGeometry;// pTopologicalOperator as IGeometry;
        }

        /// <summary>
        /// 对于少量数据(2个要素)可以用这个
        /// </summary>
        /// <param name="featureClass"></param>
        /// <returns></returns>
        public IGeometry UnionOnebyOne(IFeatureClass featureClass)
        {
            System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString() + "begin one by one");
            IFeatureCursor pFCursor = featureClass.Search(null, false);
            IFeature pFeature = pFCursor.NextFeature();    // 第一个Feature
            IGeometry pFirstGeometry = new PolygonClass();
            while (pFeature != null)
            {
                IGeometry pGeometry = pFeature.ShapeCopy;
                ITopologicalOperator pTopologicalOperator = pFirstGeometry as ITopologicalOperator;
                pFirstGeometry = pTopologicalOperator.Union(pGeometry) as IPolygon;
                System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString() + "  " + pFeature.OID);
                pFeature = pFCursor.NextFeature();
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pFCursor);
            return pFirstGeometry;
        }

        public string CheckArea(ILayer player, string pLayerCondition, TreeView treeView1, string TableName, string pTableCondition, string TableMJstring, int judge, ref DataTable pErrorDataTable, string ErrorNumber)
        {
            if (player == null) { return ""; }
            TreeNode TableFileName = CallFindNode(treeView1, TableName);
            ITable ptable = getITable(TableFileName);
            if (ptable == null) { return "\r\n" + TableName + "不存在"; }
            double area = 0;
            double area2 = 0;
            IQueryFilter pQueryFilter = new QueryFilterClass();
            pQueryFilter.WhereClause = pLayerCondition;
            IFeatureCursor pFC = (player as FeatureLayer).Search(pQueryFilter, true);
            IFeature pF = pFC.NextFeature();
            while (pF != null)
            {
                area += (pF.Shape as IArea).Area;
                pF = pFC.NextFeature();
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pFC);
            IQueryFilter pQueryFilter2 = new QueryFilterClass();
            pQueryFilter2.WhereClause = pTableCondition;
            ICursor pC = ptable.Search(pQueryFilter2, false);
            IRow pRow = pC.NextRow();
            while (pRow != null)
            {
                if (pRow.get_Value(pRow.Fields.FindField(TableMJstring)).ToString() == "") { pRow = pC.NextRow(); continue; }
                area2 += Double.Parse(pRow.get_Value(pRow.Fields.FindField(TableMJstring)).ToString());
                pRow = pC.NextRow();
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pC);
            switch (judge)
            {
                case 0:
                    if (area > area2 * 10000)
                    {
                        pErrorDataTable.Rows.Add(new object[] { ErrorNumber, player.Name, null, null, GetChineseName(player.Name) + "面积不符合与非空间数据" + TableName + "面积关系,相差面积为" + (area2 * 10000 - area) + "平方米", false, true });
                        return "\r\nERROR" + ErrorNumber + ":" + GetChineseName(player.Name) + "面积不符合与非空间数据" + TableName + "面积关系,相差面积为" + (area2 * 10000 - area) + "平方米";
                    }
                    break;
                case 1:
                    if (area < area2 * 10000)
                    {
                        pErrorDataTable.Rows.Add(new object[] { ErrorNumber, player.Name, null, null, GetChineseName(player.Name) + "面积不符合与非空间数据" + TableName + "面积关系,相差面积为" + (area - area2 * 10000) + "平方米", false, true });
                        return "\r\nERROR" + ErrorNumber + ":" + GetChineseName(player.Name) + "面积不符合与非空间数据" + TableName + "面积关系,相差面积为" + (area - area2 * 10000) + "平方米";
                    }
                    break;
                case 2:
                    if (area != area2 * 10000)
                    {
                        pErrorDataTable.Rows.Add(new object[] { ErrorNumber, player.Name, null, null, GetChineseName(player.Name) + "面积不符合与非空间数据" + TableName + "面积关系,相差面积为" + (area - area2 * 10000) + "平方米", false, true });
                        return "\r\nERROR" + ErrorNumber + ":" + GetChineseName(player.Name) + "面积不符合与非空间数据" + TableName + "面积关系,相差面积为" + (area - area2 * 10000) + "平方米";
                    }
                    break;
            }
            return "";
        }


        public string CheckArea(ILayer player, string pLayerCondition, ILayer player2, string pLayerCondition2, string TableMJstring, int judge, ref DataTable pErrorDataTable, string ErrorNumber)
        {
            if (player == null) { return ""; }
            if (player2 == null) { return ""; }
            double area = 0;
            double area2 = 0;
            IQueryFilter pQueryFilter = new QueryFilterClass();
            pQueryFilter.WhereClause = pLayerCondition;
            IFeatureCursor pFC = (player as FeatureLayer).Search(pQueryFilter, true);
            IFeature pF = pFC.NextFeature();
            while (pF != null)
            {
                area += (pF.Shape as IArea).Area;
                pF = pFC.NextFeature();
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pFC);
            IQueryFilter pQueryFilter2 = new QueryFilterClass();
            pQueryFilter2.WhereClause = pLayerCondition2;
            IFeatureCursor pC = (player2 as IFeatureLayer).Search(pQueryFilter2, true);
            IFeature pRow = pC.NextFeature();
            while (pRow != null)
            {
                area2 += Double.Parse(pRow.get_Value(pRow.Fields.FindField(TableMJstring)).ToString());
                pRow = pC.NextFeature();
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pC);
            if (pLayerCondition == null) { pLayerCondition = "总面积"; }
            if (pLayerCondition2 == null) { pLayerCondition2 = "总面积"; }
            switch (judge)
            {
                case 0:
                    if (area > area2 * 10000)
                    {
                        pErrorDataTable.Rows.Add(new object[] { ErrorNumber, player.Name, null, null, GetChineseName(player.Name) + "面积不符合与空间数据" + GetChineseName(player2.Name) + "面积关系,(即" + GetChineseName(player.Name) + " " + pLayerCondition + "<=" + GetChineseName(player2.Name) + " " + pLayerCondition + ")", false, true });
                        return "\r\nERROR" + ErrorNumber + ":" + GetChineseName(player.Name) + "面积不符合与空间数据" + GetChineseName(player2.Name) + "面积关系,(即" + GetChineseName(player.Name) + " " + pLayerCondition + "<=" + GetChineseName(player2.Name) + " " + pLayerCondition + ")";
                    }
                    break;
                case 1:
                    if (area < area2 * 10000)
                    {
                        pErrorDataTable.Rows.Add(new object[] { ErrorNumber, player.Name, null, null, GetChineseName(player.Name) + "面积不符合与空间数据" + GetChineseName(player2.Name) + "面积关系(即" + GetChineseName(player.Name) + " " + pLayerCondition + ">=" + GetChineseName(player2.Name) + " " + pLayerCondition + ")", false, true });
                        return "\r\nERROR" + ErrorNumber + ":" + GetChineseName(player.Name) + "面积不符合与空间数据" + GetChineseName(player2.Name) + "面积关系,(即" + GetChineseName(player.Name) + " " + pLayerCondition + ">=" + GetChineseName(player2.Name) + " " + pLayerCondition + ")";
                    }
                    break;
                case 2:
                    if (area != area2 * 10000)
                    {
                        pErrorDataTable.Rows.Add(new object[] { ErrorNumber, player.Name, null, null, GetChineseName(player.Name) + "面积不符合与空间数据" + GetChineseName(player2.Name) + "面积关系(即" + GetChineseName(player.Name) + " " + pLayerCondition + "=" + GetChineseName(player2.Name) + " " + pLayerCondition + "),相差面积为" + (area - area2 * 10000) + "平方米", false, true });
                        return "\r\nERROR" + ErrorNumber + ":" + GetChineseName(player.Name) + "面积不符合与空间数据" + GetChineseName(player2.Name) + "面积关系(即" + GetChineseName(player.Name) + " " + pLayerCondition + "=" + GetChineseName(player2.Name) + " " + pLayerCondition + "),相差面积为" + (area - area2 * 10000) + "平方米";
                    }
                    break;
            }
            return "";
        }


        public ITable getITable(TreeNode TN)
        {
            if (TN == null) { return null; }
            if (TN.Parent == null) { return null; }
            IWorkspaceFactory pAccessWorkspaceFactory = new AccessWorkspaceFactoryClass();
            // 打开工作空间并遍历数据集 
            IWorkspace pWorkspace = pAccessWorkspaceFactory.OpenFromFile(TN.Parent.Text, 0);
            ITable ptable = ((IFeatureWorkspace)pWorkspace).OpenTable(TN.Text);
            return ptable;
        }

        public TreeNode CallFindNode(TreeView treeView, string strValue)
        {
            TreeNodeCollection nodes = treeView.Nodes;
            foreach (TreeNode n in nodes)
            {
                TreeNode temp = FindNode(n, strValue);
                if (temp != null)
                    return temp;
            }
            return null;
        }

        public TreeNode FindNode(TreeNode tnParent, string strValue)
        {
            if (tnParent == null) return null;
            if (tnParent.Text == strValue) return tnParent;
            TreeNode tnRet = null;
            foreach (TreeNode tn in tnParent.Nodes)
            {
                tnRet = FindNode(tn, strValue);
                if (tnRet != null) break;
            }
            return tnRet;
        }

        public string CheckArea(string TableName1, string pTableCondition1, string TableMJstring, TreeView treeView1, string TableName2, string pTableCondition2, string TableMJstring2, int judge, ref DataTable pErrorDataTable, string ErrorNumber)
        {
            TreeNode TableFileName1 = CallFindNode(treeView1, TableName1);
            ITable ptable1 = getITable(TableFileName1);
            if (ptable1 == null) { return ""; }
            TreeNode TableFileName2 = CallFindNode(treeView1, TableName2);
            ITable ptable2 = getITable(TableFileName2);
            if (ptable2 == null) { return ""; }
            double area = 0;
            double area2 = 0;
            IQueryFilter pQueryFilter = new QueryFilterClass();
            pQueryFilter.WhereClause = pTableCondition1;
            ICursor pC = ptable1.Search(pQueryFilter, false);
            IRow pRow = pC.NextRow();
            while (pRow != null)
            {
                area += Double.Parse(pRow.get_Value(pRow.Fields.FindField(TableMJstring)).ToString());
                pRow = pC.NextRow();
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pC);
            IQueryFilter pQueryFilter2 = new QueryFilterClass();
            pQueryFilter2.WhereClause = pTableCondition2;
            ICursor pC2 = ptable2.Search(pQueryFilter2, false);
            IRow pRow2 = pC2.NextRow();
            while (pRow2 != null)
            {
                area2 += Double.Parse(pRow2.get_Value(pRow2.Fields.FindField(TableMJstring2)).ToString());
                pRow2 = pC2.NextRow();
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pC2);
            switch (judge)
            {
                case 0:
                    if (area > area2)
                    {
                        pErrorDataTable.Rows.Add(new object[] { ErrorNumber, TableName1, null, null, GetChineseName(TableName1) + "(" + pTableCondition1 + ")面积不大于等于与非空间数据" + GetChineseName(TableName2) + "(" + pTableCondition2 + ")面积关系,相差面积为" + (area - area2) + "公顷", false, true });
                        return "\r\nERROR" + ErrorNumber + ":" + GetChineseName(TableName1) + "(" + pTableCondition1 + ")面积不大于等于非空间数据" + GetChineseName(TableName2) + "(" + pTableCondition2 + ")面积关系,相差面积为" + (area - area2) + "公顷";
                    }
                    break;
                case 1:
                    if (area < area2)
                    {
                        pErrorDataTable.Rows.Add(new object[] { ErrorNumber, TableName1, null, null, GetChineseName(TableName1) + "(" + pTableCondition1 + ")面积不小于等于与非空间数据" + GetChineseName(TableName2) + "(" + pTableCondition2 + ")面积关系,相差面积为" + (area - area2) + "公顷", false, true });
                        return "\r\nERROR" + ErrorNumber + ":" + GetChineseName(TableName1) + "(" + pTableCondition1 + ")面积不小于等于与非空间数据" + GetChineseName(TableName2) + "(" + pTableCondition2 + ")面积关系,相差面积为" + (area - area2) + "公顷";
                    }
                    break;
                case 2:
                    if (area != area2)
                    {
                        pErrorDataTable.Rows.Add(new object[] { ErrorNumber, GetChineseName(TableName1), null, null, GetChineseName(TableName1) + "(" + pTableCondition1 + ")面积不符合与非空间数据" + GetChineseName(TableName2) + "(" + pTableCondition2 + ")面积关系,相差面积为" + (area - area2) + "公顷", false, true });
                        return "\r\nERROR" + ErrorNumber + ":" + GetChineseName(TableName1) + GetChineseName(TableName1) + "(" + pTableCondition1 + ")面积不符合与非空间数据" + GetChineseName(TableName2) + "(" + pTableCondition2 + ")面积关系,相差面积为" + (area - area2) + "公顷";
                    }
                    break;
            }
            return "";
        }

        public string CheckArea(string TableName1, string pTableCondition1, string TableMJstring, TreeView treeView1, ILayer player, string pTableCondition2, string TableMJstring2, int judge, ref DataTable pErrorDataTable, string ErrorNumber)
        {
            TreeNode TableFileName1 = CallFindNode(treeView1, TableName1);
            ITable ptable1 = getITable(TableFileName1);
            if (ptable1 == null) { return ""; }
            double area = 0;
            double area2 = 0;
            IQueryFilter pQueryFilter = new QueryFilterClass();
            pQueryFilter.WhereClause = pTableCondition1;
            ICursor pC = ptable1.Search(pQueryFilter, false);
            IRow pRow = pC.NextRow();
            while (pRow != null)
            {
                area += Double.Parse(pRow.get_Value(pRow.Fields.FindField(TableMJstring)).ToString());
                pRow = pC.NextRow();
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pC);
            IQueryFilter pQueryFilter2 = new QueryFilterClass();
            pQueryFilter2.WhereClause = pTableCondition2;
            IFeatureCursor pC2 = (player as IFeatureLayer).FeatureClass.Search(pQueryFilter2, false);
            IFeature pRow2 = pC2.NextFeature(); 
            while (pRow2 != null)
            {
                area2 += Double.Parse(pRow2.get_Value(pRow2.Fields.FindField(TableMJstring2)).ToString());
                pRow2 = pC2.NextFeature(); 
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pC2);
            switch (judge)
            {
                case 0:
                    if (area > area2)
                    {
                        pErrorDataTable.Rows.Add(new object[] { ErrorNumber, TableName1, null, null, GetChineseName(TableName1) + "面积不符合与空间数据" + GetChineseName(player.Name) + "面积关系", false, true });
                        return "\r\nERROR" + ErrorNumber + ":非空间数据" + GetChineseName(TableName1) + "面积不符合与非空间数据" + GetChineseName(player.Name) + "面积关系";
                    }
                    break;
                case 1:
                    if (area < area2)
                    {
                        pErrorDataTable.Rows.Add(new object[] { ErrorNumber, TableName1, null, null, GetChineseName(TableName1) + "面积不符合与空间数据" + GetChineseName(player.Name) + "面积关系", false, true });
                        return "\r\nERROR" + ErrorNumber + ":非空间数据" + GetChineseName(TableName1) + "面积不符合与非空间数据" + GetChineseName(player.Name) + "面积关系";
                    }
                    break;
                case 2:
                    if (area != area2)
                    {
                        pErrorDataTable.Rows.Add(new object[] { ErrorNumber, GetChineseName(TableName1), null, null, "面积不符合与空间数据" + GetChineseName(player.Name) + "面积关系,相差面积为" + (area - area2) + "公顷", false, true });
                        return "\r\nERROR" + ErrorNumber + ":" + GetChineseName(TableName1) + "面积不符合与空间数据" + GetChineseName(player.Name) + "面积关系,相差面积为" + (area - area2) + "公顷";
                    }
                    break;
            }
            return "";
        }




        public double getArea(string TableName, string pTableCondition, string TableMJstring, TreeView treeView1)
        {
            TreeNode TableFileName = CallFindNode(treeView1, TableName);
            ITable ptable1 = getITable(TableFileName);
            if (ptable1 == null) { return -999; }
            double area = 0;
            IQueryFilter pQueryFilter = new QueryFilterClass();
            pQueryFilter.WhereClause = pTableCondition;
            ICursor pC = ptable1.Search(pQueryFilter, false);
            IRow pRow = pC.NextRow();
            while (pRow != null)
            {
                area += Double.Parse(pRow.get_Value(pRow.Fields.FindField(TableMJstring)).ToString());
                pRow = pC.NextRow();
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pC);
            return area;
        }

        public string JudgeArea(double area1, string TableName, double area2, string TableName2, int judge, ref DataTable pErrorDataTable, string ErrorNumber)
        {
            if (area1 == -999) { return "\r\n" + TableName + "不存在"; }
            if (area2 == -999) { return "\r\n" + TableName2 + "不存在"; }
            switch (judge)
            {
                case 0:
                    if (Math.Round(area1, 2) > Math.Round(area2, 2))
                    {
                        pErrorDataTable.Rows.Add(new object[] { ErrorNumber, TableName, null, null, GetChineseName(TableName)+"面积不符合与非空间数据" + GetChineseName(TableName2) + "面积关系", false, true });
                        return "\r\nERROR" + ErrorNumber + ":非空间数据" + TableName + "面积不符合与非空间数据" + GetChineseName(TableName2) + "面积关系";
                    }
                    break;
                case 1:
                    if (Math.Round(area1, 2) < Math.Round(area2, 2))
                    {
                        pErrorDataTable.Rows.Add(new object[] { ErrorNumber, TableName, null, null, GetChineseName(TableName)+"面积不符合与非空间数据" + GetChineseName(TableName2) + "面积关系", false, true });
                        return "\r\nERROR" + ErrorNumber + ":非空间数据" + GetChineseName(TableName) + "面积不符合与非空间数据" + GetChineseName(TableName2) + "面积关系";
                    }
                    break;
                case 2:
                    if (Math.Round(area1, 2) != Math.Round(area2, 2))
                    {
                        pErrorDataTable.Rows.Add(new object[] { ErrorNumber, TableName, null, null,GetChineseName(TableName)+"面积不符合与非空间数据" + GetChineseName(TableName2) + "面积关系,相差面积为" + (area1 - area2) + "公顷", false, true });
                        return "\r\nERROR" + ErrorNumber + ":" + GetChineseName(TableName) + "面积不符合与非空间数据" + GetChineseName(TableName2) + "面积关系,相差面积为" + (area1 - area2) + "公顷";
                    }
                    break;
            }
            return "";
        }

        public string CheckLayerAttribute(ILayer player, ref DataTable pErrorDataTable, string condition1, string LayerAttrName2, string datarange)
        {
            string Error = "";
            if (player == null) { return ""; }
            IQueryFilter pQueryFilter = new QueryFilterClass();
            pQueryFilter.WhereClause = condition1;
            IFeatureCursor pFCursor = (player as FeatureLayer).FeatureClass.Search(pQueryFilter, false);
            IFeature pFeature = pFCursor.NextFeature();
            while (pFeature != null)
            {
                string value = pFeature.get_Value(pFeature.Fields.FindField(LayerAttrName2)).ToString();
                if (value[0] != 1 || value[1] != 1)
                {
                    Error += "\r\nERROR6503:图层" + GetChineseName(player.Name) + "的要素OID：" + pFeature.OID + "管控等级与保护类型代码不匹配";
                    pErrorDataTable.Rows.Add(new object[] { "6503", player.Name, null, pFeature.OID.ToString(), GetChineseName(player.Name) + "的要素OID：" + pFeature.OID + "管控等级与保护类型代码不匹配", false, true });
                }
                pFeature = pFCursor.NextFeature();
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pFCursor);
            return Error;
        }

        public double[] StatisticalScaleGHFW(ILayer player)
        {
            string[] AttrsEng ={"GDBYL","YJJBNTBHRW","XZJSYDZGDMJ","TDZZBCGD","XZJSYDJLHMJ","STBHHXMJ",
                                   "STKJMJ","CSKFBJMJ","JSYDZGM","CSKFBJNXZJSYDMJ"};
            double[] areas = new double[10]{0,0,0,0,0,0,0,0,0,0};
            if (player == null) { return areas; }
            IFeatureCursor pFC=(player as IFeatureLayer).Search(null, true);
            IFeature pFeature = pFC.NextFeature();
            while (pFeature != null)
            {
                for(int i=0;i<AttrsEng.Length;i++)
                {
                    int OIDindex = pFeature.Fields.FindField(AttrsEng[i]);
                    areas[i] += double.Parse(pFeature.get_Value(OIDindex).ToString());
                }
                pFeature = pFC.NextFeature();
            }
            return areas;
        }

        public double getLayerArea(ILayer player,string Condition)
        {
            if (player == null) { return 0; }
            double area = 0;
            IQueryFilter pQF = new QueryFilterClass();
            pQF.WhereClause = Condition;
            IFeatureCursor pFC = (player as IFeatureLayer).Search(pQF, true);
            IFeature pF = pFC.NextFeature();
            while (pF != null)
            {
                area += (pF.Shape as IArea).Area;
                pF = pFC.NextFeature();
            }
            return area/10000;
        }

        public double getUnionArea(ILayer player1, ILayer player2, string Condition1, string Condition2)
        {
            if (player1 == null || player2 == null) { return 0; }
            ISpatialReference GRout = ((player1 as FeatureLayer).FeatureClass as IGeoDataset).SpatialReference;
            ISpatialReference GRin = ((player2 as FeatureLayer).FeatureClass as IGeoDataset).SpatialReference;
            if (GRout.SpatialReferenceImpl != GRin.SpatialReferenceImpl || GRout.Name != GRin.Name)
            { MessageBox.Show(player1.Name + "与" + player2.Name + "的坐标系不同，" + player2.Name + "与" + player1.Name + "范围无法比较，请修改"); return 0; }
            //string relationDescription = "RELATE(G1, G2, '**T******')";
            IQueryFilter pQueryFilter = new QueryFilterClass();
            pQueryFilter.WhereClause = Condition1;
            IGeometry Geo1 = UnionAllSelect((player1 as FeatureLayer).FeatureClass, pQueryFilter);
            IQueryFilter pQueryFilter2 = new QueryFilterClass();
            pQueryFilter.WhereClause = Condition2;
            IGeometry Geo2 = UnionAllSelect((player2 as FeatureLayer).FeatureClass, pQueryFilter2);

            ITopologicalOperator pGeoInTP = Geo1 as ITopologicalOperator;
            IGeometry pUnion = pGeoInTP.Union(Geo2);
            IArea pArea = pUnion as IArea;
            if (Geo2 != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Geo2);
            if (pGeoInTP != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pGeoInTP);
            return pArea.Area / 10000;
        }
        public double getIntersectArea(ILayer player1, ILayer player2, string Condition1, string Condition2)
        {
            if (player1 == null || player2 == null) { return 0; }
            ISpatialReference GRout = ((player1 as FeatureLayer).FeatureClass as IGeoDataset).SpatialReference;
            ISpatialReference GRin = ((player2 as FeatureLayer).FeatureClass as IGeoDataset).SpatialReference;
            if (GRout.SpatialReferenceImpl != GRin.SpatialReferenceImpl || GRout.Name != GRin.Name)
            { MessageBox.Show(player1.Name + "与" + player2.Name + "的坐标系不同，" + player2.Name + "与" + player1.Name + "范围无法比较，请修改"); return 0; }
            //string relationDescription = "RELATE(G1, G2, '**T******')";
            IQueryFilter pQueryFilter = new QueryFilterClass();
            pQueryFilter.WhereClause = Condition1;
            IGeometry Geo1 = UnionAllSelect((player1 as FeatureLayer).FeatureClass, pQueryFilter);
            double pareaGeo1 = (Geo1 as IArea).Area;
            IQueryFilter pQueryFilter2 = new QueryFilterClass();
            pQueryFilter2.WhereClause = Condition2;
            IGeometry Geo2 = UnionAllSelect((player2 as FeatureLayer).FeatureClass, pQueryFilter2);
            double pareaGeo2 = (Geo2 as IArea).Area;

            ITopologicalOperator pGeoInTP = Geo1 as ITopologicalOperator;
            IGeometry pIntersect = pGeoInTP.Intersect(Geo2,esriGeometryDimension.esriGeometry2Dimension);
            IArea pArea = pIntersect as IArea;
            if (Geo2 != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Geo2);
            if (pGeoInTP != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pGeoInTP);
            return pArea.Area / 10000;
        }

        public string StatisticalScale(ILayer player1, ILayer player2, ref DataTable pErrorDataTable,string TableName,TreeView treeView1)
        {
            if (player1 == null||player2 ==null) { return "\r\n图层JQDLTB或JSYDHJBNTGZ2035不存在"; }
            ISpatialReference GRout = ((player1 as FeatureLayer).FeatureClass as IGeoDataset).SpatialReference;
            ISpatialReference GRin = ((player2 as FeatureLayer).FeatureClass as IGeoDataset).SpatialReference;
            if (GRout.SpatialReferenceImpl != GRin.SpatialReferenceImpl || GRout.Name != GRin.Name)
            { MessageBox.Show("外部数据" + player1.Name + "与" + player2.Name + "的坐标系不同，" + player2.Name + "与" + player1.Name + "范围无法比较，请修改"); return "\r\n外部数据" + player1.Name + "与" + player2.Name + "的坐标系不同，" + player2.Name + "与" + player1.Name + "范围无法比较，请修改"; }

            //string relationDescription = "RELATE(G1, G2, '**T******')";
            IQueryFilter pQueryFilter = new QueryFilterClass();
            pQueryFilter.WhereClause = "DLBM_SX LIKE '2*'";
            IGeometry Geo1 = UnionAllSelect((player1 as FeatureLayer).FeatureClass, pQueryFilter);
            IQueryFilter pQueryFilter2 = new QueryFilterClass();
            pQueryFilter.WhereClause = "GZQLXDM LIKE '011' OR GZQLXDM LIKE '012'";
            IGeometry Geo2 = UnionAllSelect((player2 as FeatureLayer).FeatureClass, pQueryFilter2);
            double areaB=(Geo2 as IArea).Area;

            ITopologicalOperator pGeoInTP = Geo1 as ITopologicalOperator;
            IGeometry pDiff = pGeoInTP.Difference(Geo2);
            IArea pArea = pDiff as IArea;
            //IRelationalOperator RO = Geo2 as IRelationalOperator;
            //bool isEqual = RO.Within(Geo1);//.Relation(Geo1, relationDescription);
            if (Geo2 != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Geo2);
            if (pGeoInTP != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pGeoInTP);
            System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString() + " end diff");
            TreeNode TableFileName = CallFindNode(treeView1, TableName);
            ITable ptable = getITable(TableFileName);
            if (ptable == null) { return "\r\n非空间数据表"+TableName+"不存在"; }
            IQueryFilter pQueryFilter3 = new QueryFilterClass();
            pQueryFilter.WhereClause = "ZBDM LIKE '06'";
            ICursor pCursor = ptable.Search(pQueryFilter3, true);
            IRow pRow = pCursor.NextRow();
            double areaC = 0;
            while(pRow!=null)
            {
                double arearow = Double.Parse(pRow.get_Value(pRow.Fields.FindField("ZBMJ")).ToString());
                areaC += arearow;
                pRow = pCursor.NextRow();
            }
            return "\r\n建设用地减量化规模为"+((pArea.Area+areaB-areaC*10000)/10000).ToString()+"公顷";
        }

        public DataTable GetExcelTable(string excelFilename)
        {
            string connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Jet OLEDB:Engine Type=35;Extended Properties=Excel 8.0;Persist Security Info=False", excelFilename);
            //string connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=Yes;IMEX=1;'", excelFilename);
            DataSet ds = new DataSet();
            string tableName;
            using (System.Data.OleDb.OleDbConnection connection = new System.Data.OleDb.OleDbConnection(connectionString))
            {
                connection.Open();
                DataTable table = connection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
                tableName = table.Rows[0]["Table_Name"].ToString();
                string strExcel = "select * from " + "[" + tableName + "]";
                OleDbDataAdapter adapter = new OleDbDataAdapter(strExcel, connectionString);
                adapter.Fill(ds, tableName);
                connection.Close();
            }
            return ds.Tables[tableName];
        }

        public string CheckTableAndExcelArea(string TableName1, string pTableCondition1, string TableMJstring, TreeView treeView1, string ExcelName, string pExcelCondition2, int judge, ref DataTable pErrorDataTable, string ErrorNumber,int ExcelColumnUnitIndex,string ExcelColumn)
        {
            DataTable DT = GetExcelTable(ExcelName);
            TreeNode TableFileName1 = CallFindNode(treeView1, TableName1);
            ITable ptable1 = getITable(TableFileName1);
            if (ptable1 == null) { return ""; }
            double area = 0;
            double area2 = 0;
            IQueryFilter pQueryFilter = new QueryFilterClass();
            pQueryFilter.WhereClause = pTableCondition1;
            ICursor pC = ptable1.Search(pQueryFilter, false);
            IRow pRow = pC.NextRow();
            while (pRow != null)
            {
                area += Double.Parse(pRow.get_Value(pRow.Fields.FindField(TableMJstring)).ToString());
                pRow = pC.NextRow();
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pC);
            DataRow[] DRs = DT.Select(pExcelCondition2);
            for (int i=0;i<DRs.Length ;i++)
            {
                area2 += Double.Parse(DRs[i][ExcelColumn].ToString());
            }
            double unit = 0;
            switch (ExcelColumnUnitIndex)
            {
                case 1:
                case 2:
                    unit = 666.67;
                    break;
                case 3:
                    unit = 100;
                    break;
                default:
                    unit = 0;
                    break;
            }
            switch (judge)
            {
                case 0:
                    if (area > area2 * unit)
                    {
                        pErrorDataTable.Rows.Add(new object[] { ErrorNumber, TableName1, null, null, GetChineseName(TableName1) + "面积不符合与外部数据面积关系,相差面积为" + (area - area2 * unit) + "平方米", false, true });
                        return "\r\nERROR" + ErrorNumber + ":非空间数据" + GetChineseName(TableName1) + "面积不符合与外部数据面积关系,相差面积为" + (area - area2 * unit) + "平方米";
                    }
                    break;
                case 1:
                    if (area < area2 * unit)
                    {
                        pErrorDataTable.Rows.Add(new object[] { ErrorNumber, TableName1, null, null, GetChineseName(TableName1) + "面积不符合与外部数据面积关系,相差面积为" + (area - area2 * unit) + "平方米", false, true });
                        return "\r\nERROR" + ErrorNumber + ":非空间数据" + GetChineseName(TableName1) + "面积不符合与外部数据面积关系,相差面积为" + (area - area2 * unit) + "平方米";
                    }
                    break;
                case 2:
                    if (area != area2 * unit)
                    {
                        pErrorDataTable.Rows.Add(new object[] { ErrorNumber, GetChineseName(TableName1), null, null, "面积不符合与外部数据面积关系,相差面积为" + (area - area2 * unit) + "公顷", false, true });
                        return "\r\nERROR" + ErrorNumber + ":" + GetChineseName(TableName1) + "面积不符合与外部数据面积关系,相差面积为" + (area - area2 * unit) + "平方米";
                    }
                    break;
            }
            return "";
        }


        public DataTable getLayerAreaByCity_1(ILayer player, string Condition,ILayer City)
        {
            if (player == null) { return null; }
            DataTable dt = new DataTable();
            DataColumn dc1 = new DataColumn("镇名称", Type.GetType("System.String"));
            DataColumn dc2 = new DataColumn("01*+02*面积", Type.GetType("System.String"));
            dt.Columns.Add(dc1);
            dt.Columns.Add(dc2);
            IQueryFilter pQF=new QueryFilterClass();
            pQF.WhereClause = Condition;
            IGeometry GeoLayer = UnionAllSelect((player as IFeatureLayer).FeatureClass, pQF);
            ITopologicalOperator ITO = GeoLayer as ITopologicalOperator;
            IFeatureCursor pFC = (City as IFeatureLayer).Search(null, true);
            IFeature pF = pFC.NextFeature();
            while (pF != null)
            {
                IGeometry pIntersect = ITO.Intersect(pF.Shape, esriGeometryDimension.esriGeometry2Dimension);
                DataRow dr = dt.NewRow();
                dr["镇名称"] = pF.get_Value(pF.Fields.FindField("XZBJMC")).ToString();
                dr["01*+02*面积"] = (pIntersect as IArea).Area.ToString();
                dt.Rows.Add(dr);
                pF = pFC.NextFeature();
            }
            return dt;
        }

        public DataTable getLayerAreaByCity_2(ILayer player, string Condition,string Condition2, ILayer City)
        {
            if (player == null) { return null; }
            DataTable dt = new DataTable();
            DataColumn dc1 = new DataColumn("镇名称", Type.GetType("System.String"));
            DataColumn dc2 = new DataColumn("城市开发边界规模", Type.GetType("System.String"));
            DataColumn dc3 = new DataColumn("城市开发边界内建设用地规模", Type.GetType("System.String"));
            dt.Columns.Add(dc1);
            dt.Columns.Add(dc2); 
            dt.Columns.Add(dc3);
            IQueryFilter pQF1 = new QueryFilterClass();
            pQF1.WhereClause = Condition;
            IQueryFilter pQF2 = new QueryFilterClass();
            pQF2.WhereClause = Condition2;
            IGeometry GeoLayer = UnionAllSelect((player as IFeatureLayer).FeatureClass, pQF1);
            IGeometry GeoLayer2 = UnionAllSelect((player as IFeatureLayer).FeatureClass, pQF2);
            ITopologicalOperator ITO = GeoLayer as ITopologicalOperator;
            ITopologicalOperator ITO2 = GeoLayer2 as ITopologicalOperator;
            IFeatureCursor pFC = (City as IFeatureLayer).Search(null, true);
            IFeature pF = pFC.NextFeature();
            while (pF != null)
            {
                IGeometry pIntersect1 = ITO.Intersect(pF.Shape, esriGeometryDimension.esriGeometry2Dimension);
                IGeometry pIntersect2 = ITO2.Intersect(pF.Shape, esriGeometryDimension.esriGeometry2Dimension);
                DataRow dr = dt.NewRow();
                dr["镇名称"] = pF.get_Value(pF.Fields.FindField("XZBJMC")).ToString();
                dr["城市开发边界规模"] = (pIntersect1 as IArea).Area.ToString();
                dr["城市开发边界内建设用地规模"] = (pIntersect2 as IArea).Area.ToString();
                dt.Rows.Add(dr);
                pF = pFC.NextFeature();
            }
            return dt;
        }

        public DataTable getLayerAreaByCity_3(ILayer JSYDH, ILayer STBH, string Condition, string Condition2,string Condition3, ILayer City)
        {
            if (JSYDH == null || STBH == null || City==null) { return null; }
            DataTable dt = new DataTable();
            DataColumn dc1 = new DataColumn("镇名称", Type.GetType("System.String"));
            DataColumn dc2 = new DataColumn("生态空间面积01", Type.GetType("System.String"));
            DataColumn dc3 = new DataColumn("生态空间面积02", Type.GetType("System.String"));
            dt.Columns.Add(dc1);
            dt.Columns.Add(dc2);
            dt.Columns.Add(dc3);
            IQueryFilter pQF1 = new QueryFilterClass();
            pQF1.WhereClause = Condition;
            IQueryFilter pQF2 = new QueryFilterClass();
            pQF2.WhereClause = Condition2;
            IQueryFilter pQF3 = new QueryFilterClass();
            pQF3.WhereClause = Condition3;
            IGeometry GeoLayer = UnionAllSelect((JSYDH as IFeatureLayer).FeatureClass, pQF1);
            IGeometry GeoLayer2 = UnionAllSelect((STBH as IFeatureLayer).FeatureClass, pQF2);
            IGeometry GeoLayer3 = UnionAllSelect((STBH as IFeatureLayer).FeatureClass, pQF3);
            ITopologicalOperator ITO = GeoLayer as ITopologicalOperator;
            IGeometry Intersect2 =ITO.Intersect(GeoLayer2, esriGeometryDimension.esriGeometry2Dimension);
            IGeometry Intersect3 = ITO.Intersect(GeoLayer3, esriGeometryDimension.esriGeometry2Dimension);
            IFeatureCursor pFC = (City as IFeatureLayer).Search(null, true);
            IFeature pF = pFC.NextFeature();
            ITopologicalOperator ITOIntersect2 = Intersect2 as ITopologicalOperator;
            ITopologicalOperator ITOIntersect3 = Intersect3 as ITopologicalOperator;
            while (pF != null)
            {
                IGeometry pIntersect2 = ITOIntersect2.Intersect(pF.Shape, esriGeometryDimension.esriGeometry2Dimension);
                IGeometry pIntersect3 = ITOIntersect3.Intersect(pF.Shape, esriGeometryDimension.esriGeometry2Dimension);
                DataRow dr = dt.NewRow();
                dr["镇名称"] = pF.get_Value(pF.Fields.FindField("XZBJMC")).ToString();
                dr["生态空间面积01"] = (pIntersect2 as IArea).Area.ToString();
                dr["生态空间面积02"] = (pIntersect3 as IArea).Area.ToString();
                dt.Rows.Add(dr);
                pF = pFC.NextFeature();
            }
            return dt;
        }

        public DataTable getLayerAreaByCity_4(ILayer JQDLTB, ILayer JSYDH, string Condition, string Condition2, ILayer City)// ,ITable ZBFJB,TreeView treeview1)
        {
            if (JSYDH == null || JQDLTB == null || City == null) { return null; }
            DataTable dt = new DataTable();
            DataColumn dc1 = new DataColumn("镇名称", Type.GetType("System.String"));
            DataColumn dc2 = new DataColumn("现状建设用地减量化", Type.GetType("System.String"));
            dt.Columns.Add(dc1);
            dt.Columns.Add(dc2);
            IQueryFilter pQF1 = new QueryFilterClass();
            pQF1.WhereClause = Condition;
            IQueryFilter pQF2 = new QueryFilterClass();
            pQF2.WhereClause = Condition2;
            IGeometry GeoLayer = UnionAllSelect((JQDLTB as IFeatureLayer).FeatureClass, pQF1);
            IGeometry GeoLayer2 = UnionAllSelect((JSYDH as IFeatureLayer).FeatureClass, pQF2);
            ITopologicalOperator ITO = GeoLayer as ITopologicalOperator;
            IGeometry union = ITO.Union(GeoLayer2);
            IFeatureCursor pFC = (City as IFeatureLayer).Search(null, true);
            IFeature pF = pFC.NextFeature();
            ITopologicalOperator ITOunion = union as ITopologicalOperator;
            while (pF != null)
            {
                IGeometry pIntersect = ITOunion.Intersect(pF.Shape, esriGeometryDimension.esriGeometry2Dimension);
                DataRow dr = dt.NewRow();
                dr["镇名称"] = pF.get_Value(pF.Fields.FindField("XZBJMC")).ToString();
                dr["现状建设用地减量化"] = (pIntersect as IArea).Area.ToString();
                dt.Rows.Add(dr);
                pF = pFC.NextFeature();
            }
            return dt;
        }
    }
}
