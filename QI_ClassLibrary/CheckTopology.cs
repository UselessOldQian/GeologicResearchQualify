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
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Controls;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.DataSourcesFile;
using ESRI.ArcGIS.DataSourcesGDB;
using ESRI.ArcGIS.ADF;
using ESRI.ArcGIS.esriSystem;

namespace QI_ClassLibrary
{
    public class CheckTopology
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
        /// 检查数据是否自相交
        /// </summary>
        /// <param name="player"></param>
        /// <param name="pErrorDataTable"></param>
        /// <returns></returns>
        public string CheckSelfIntersection(ILayer player, ref DataTable pErrorDataTable)
        {
            if (player == null) { return ""; }
            string Error = "";
            IFeature pFeature;
            IFeatureCursor pFCursor = (player as FeatureLayer).FeatureClass.Search(null, false);
            pFeature = pFCursor.NextFeature();    // 第一个Feature
            while (pFeature != null) // 如果是null就说明刚才那个是最后一个了，后面没有了
            {
                ITopologicalOperator3 pTopologicalOperator3 = pFeature.Shape as ITopologicalOperator3;
                pTopologicalOperator3.IsKnownSimple_2 = false;
                esriNonSimpleReasonEnum reason = esriNonSimpleReasonEnum.esriNonSimpleOK;
                if (!pTopologicalOperator3.get_IsSimpleEx(out reason))
                {
                    if (reason == esriNonSimpleReasonEnum.esriNonSimpleSelfIntersections)
                    {
                        Error += "\r\nERROR4101:" + GetChineseName(player.Name) + "的ID=" + pFeature.OID.ToString() + "要素自相交";
                        pErrorDataTable.Rows.Add(new object[] { "4101", player.Name, null, pFeature.OID.ToString(), GetChineseName(player.Name) + "的ID=" + pFeature.OID.ToString() + "要素自相交", false, true });
                    }
                }
                // 对pFeature作相应处理
                pFeature = pFCursor.NextFeature();    // 下一个Feature
            }
            return Error;
        }


        /// <summary>
        /// 检查图层内要素是否有重叠,由于JSYDHJBNTGZ2035检查很慢，所以FeaturesOnetime在5左右就行
        /// </summary>
        /// <param name="player"></param>
        /// <param name="pErrorDataTable"></param>
        /// <returns></returns>
        public string CheckSelfOverlap(ILayer player, ref DataTable pErrorDataTable, int FeaturesOnetime)
        {
            string Error = "";
            try
            {
                if (player == null) { return ""; }
                IFeature pFeature;
                IFeatureClass pFeatureClass = (player as FeatureLayer).FeatureClass;
                IFeatureClass pNewMemoryFeatureclass;
                IFeatureCursor pFCursor = pFeatureClass.Search(null, false);
                pFeature = pFCursor.NextFeature();    // 第一个Feature

                int count = pFeatureClass.FeatureCount(null);

                IGeometryBag pGeometryBag = new GeometryBag() as IGeometryBag;
                pGeometryBag.SpatialReference = (pFeatureClass as IGeoDataset).SpatialReference;
                IGeometryCollection pGeometryCollection = pGeometryBag as IGeometryCollection;
                object missing = Type.Missing;
                List<IFeature> featureList = new List<IFeature>();
                int featureNum = 0;
                while (pFeature != null) // 如果是null就说明刚才那个是最后一个了，后面没有了
                {
                    featureNum++;
                    featureList.Add(pFeature);
                    pGeometryCollection.AddGeometry(pFeature.ShapeCopy, missing, missing);

                    if (featureNum % FeaturesOnetime == 0 || featureNum == count)
                    {
                        //ISpatialIndex pSpatialIndex = pGeometryBag as ISpatialIndex;
                        //pSpatialIndex.AllowIndexing = true;
                        //pSpatialIndex.Invalidate();
                        pNewMemoryFeatureclass = CreateMemoryFeatureClass(pFeatureClass);
                        IFeatureCursor pInsertCurosr = pNewMemoryFeatureclass.Insert(true);
                        IFeatureBuffer pFeatureBuffer = pNewMemoryFeatureclass.CreateFeatureBuffer();

                        ISpatialFilter pSpatialFilter1 = new SpatialFilterClass();
                        pSpatialFilter1.Geometry = pGeometryCollection as IGeometry;
                        pSpatialFilter1.SpatialRel = esriSpatialRelEnum.esriSpatialRelOverlaps;

                        IFeatureCursor pFC50 = pFeatureClass.Search(pSpatialFilter1, false);
                        IFeature pFeature50 = pFC50.NextFeature();
                        while (pFeature50 != null)
                        {
                            //IFeature CreatedFeature=pNewMemoryFeatureclass.CreateFeature();
                            //CreatedFeature = pFeature50;
                            //CreatedFeature.Store();
                            pFeatureBuffer.Shape = pFeature50.ShapeCopy;
                            pInsertCurosr.InsertFeature(pFeatureBuffer);
                            pFeature50 = pFC50.NextFeature();
                        }
                        pInsertCurosr.Flush();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(pFC50);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(pInsertCurosr);
                        if (pNewMemoryFeatureclass.FeatureCount(null) != 0)
                        {
                            System.Diagnostics.Debug.WriteLine(pNewMemoryFeatureclass.FeatureCount(null));
                            for (int j = 0; j < featureList.Count; j++)
                            {
                                IFeature pNewFeature = featureList[j];
                                ESRI.ArcGIS.Geodatabase.ISpatialFilter pSpatialFilter = new ESRI.ArcGIS.Geodatabase.SpatialFilterClass();
                                pSpatialFilter.SpatialRel = ESRI.ArcGIS.Geodatabase.esriSpatialRelEnum.esriSpatialRelOverlaps;
                                pSpatialFilter.Geometry = pNewFeature.ShapeCopy;
                                IFeatureCursor pFCursor2 = pFeatureClass.Search(pSpatialFilter, false);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(pSpatialFilter);
                                IFeature pFeature2 = pFCursor2.NextFeature();
                                System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString() + "  " + pNewFeature.OID);
                                while (pFeature2 != null)
                                {
                                    Error += "\r\nERROR4101:" + GetChineseName(player.Name) + "的ID=" + pNewFeature.OID.ToString() + "与ID=" + pFeature2.OID.ToString() + "重叠";
                                    pErrorDataTable.Rows.Add(new object[] { "4101", player.Name, null, pNewFeature.OID.ToString(), GetChineseName(player.Name) + "的ID=" + pNewFeature.OID.ToString() + "与ID=" + pFeature2.OID.ToString() + "重叠", false, true });
                                    pFeature2 = pFCursor2.NextFeature();
                                }
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(pFCursor2);
                            }
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(pNewMemoryFeatureclass);
                        featureList.Clear();
                        pGeometryCollection.RemoveGeometries(0, pGeometryCollection.GeometryCount);

                    }
                    //ESRI.ArcGIS.ADF.ComReleaser.ReleaseCOMObject(pFCursor2);
                    System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString() + "  " + pFeature.OID);
                    // 对pFeature作相应处理
                    pFeature = pFCursor.NextFeature();    // 下一个Feature
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(pFCursor);
                return Error;
            }
            catch (Exception ex)
            {
                MessageBox.Show("拓扑检查出错!请检查 | " + ex.Message);
                return Error;//出错则不返回
            }
        }

        public IFeatureClass CreateMemoryFeatureClass(IFeatureClass pFeatureClass)
        {
            IWorkspaceFactory pWSF = new InMemoryWorkspaceFactoryClass();
            IWorkspaceName pWSName = pWSF.Create("", "Temp", null, 0);
            IName pName = (IName)pWSName;
            IWorkspace memoryWS = (IWorkspace)pName.Open();

            //创建要素类            
            IFeatureWorkspace featureWorkspace = (IFeatureWorkspace)memoryWS;
            IFeatureClass featureClass = featureWorkspace.CreateFeatureClass("tempFeatureClass", pFeatureClass.Fields, null, null, pFeatureClass.FeatureType, pFeatureClass.ShapeFieldName, "");
            return featureClass;

        }

        /// <summary>
        /// 检查要素是否闭合
        /// </summary>
        /// <param name="player"></param>
        /// <param name="pErrorDataTable"></param>
        /// <returns></returns>
        public string CheckSimple(ILayer player, ref DataTable pErrorDataTable)
        {
            if (player == null) { return ""; }
            string Error = "";
            IFeature pFeature;
            IFeatureCursor pFCursor = (player as FeatureLayer).FeatureClass.Search(null, false);
            pFeature = pFCursor.NextFeature();    // 第一个Feature
            while (pFeature != null) // 如果是null就说明刚才那个是最后一个了，后面没有了
            {
                ITopologicalOperator pTopologBoundary = pFeature.Shape as ITopologicalOperator;
                bool bIsSimple = pTopologBoundary.IsSimple;
                if (bIsSimple == false)
                {
                    Error += "\r\nERROR4101:" + GetChineseName(player.Name) + "的ID=" + pFeature.OID.ToString() + "不闭合";
                    pErrorDataTable.Rows.Add(new object[] { "4101", player.Name, null, pFeature.OID.ToString(), GetChineseName(player.Name) + "的ID=" + pFeature.OID.ToString() + "不闭合", false, true });
                }
                // 对pFeature作相应处理
                pFeature = pFCursor.NextFeature();    // 下一个Feature
            }
            return Error;
        }

    }
}
