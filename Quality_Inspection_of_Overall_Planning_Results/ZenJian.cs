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
using System.Collections; 
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
using QI_ClassLibrary;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Globalization;

namespace Quality_Inspection_of_Overall_Planning_Results
{
    public partial class ZenJian : Form
    {

        private frmLoading loadForm;
        MapAction _mapAction = MapAction.Null;
        IFeatureWorkspace pFeatureWorkspace;
        IFeatureLayer pFeatureLayer;
        IFeatureDataset pFeatureDataset;
        ILayer selectedLayer;
        public DataTable pDT;

        public delegate void AppendTextInfo(string strMsg);
        public AppendTextInfo myDelegateAppendTextInfo;

        public delegate void UpdateBarValue(int iValue);
        public UpdateBarValue myDelegateUpdateBarValue;

        public delegate void UpdateUiStatus(string strMsg);
        public UpdateUiStatus myDelegateUpdateUiStatus;

        public int ProcessBarMaxValue = 0;
        public bool IsRun = false;
        List<String> FieldName = new List<string>();
        public ZenJian()
        {
            //创建加载窗体             
            loadForm = new frmLoading();
            //指定窗体加载完毕时的事件
            this.Shown += FrmLoading_Close;
            loadForm.Show();
            //主窗体初始化方法
            InitializeComponent();
            Initialize();
            axTOCControl1.SetBuddyControl(axMapControl1);
            axTOCControl1.EnableLayerDragDrop = true;


            //设置表格背景色
            dgvTable.RowsDefaultCellStyle.BackColor = Color.Ivory;

            //设置交替行的背景色
            dgvTable.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;
        }

        private void Initialize()
        {
            myDelegateAppendTextInfo = new AppendTextInfo(AppendTextInfoMethod);
            myDelegateUpdateBarValue = new UpdateBarValue(UpdateBarValueMethod);
            myDelegateUpdateUiStatus = new UpdateUiStatus(UpdateStatusBarMethod);
        }

        public void AppendTextInfoMethod(string strMsg)
        {
            //新增注释
            //if (null != InformationBox && !InformationBox.IsDisposed && strMsg != null)
            //{
            //    InformationBox.AppendText(strMsg);
            //}
        }

        public void UpdateBarValueMethod(int iValue)
        {
            if (null != progressBar1 && !progressBar1.IsDisposed)
            {
                progressBar1.Value = iValue;
            }
        }

        public void UpdateStatusBarMethod(string strMsg)
        {
            if (null != uiStatusBar1 && !uiStatusBar1.IsDisposed && strMsg != null)
            {
                uiStatusBar1.Panels[0].Text = strMsg;
            }
        }





        //声明关闭加载窗体方法
        private void FrmLoading_Close(object sender, EventArgs e)
        {
            loadForm.Close();
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {
            Thread.Sleep(2000);
        }

        private void btnLoadMDB_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            try
            {
                System.Windows.Forms.OpenFileDialog openShipFileDlg = new System.Windows.Forms.OpenFileDialog();
                openShipFileDlg.Filter = "MDB文件(*.mdb)|*.mdb";
                openShipFileDlg.Multiselect = false;
                openShipFileDlg.Title = "选择MDB文件";
                openShipFileDlg.RestoreDirectory = true;
                DialogResult dr = openShipFileDlg.ShowDialog();
                if (dr == DialogResult.OK)
                {
                    string strFullPath = openShipFileDlg.FileName;
                    if (strFullPath == "") return;
                    OpenMDB(strFullPath);
                    this.tabMapTableView.SelectedTab = tabMapTableView.TabPages[0];
                    uiStatusBar1.Panels[0].Text = "数据库读取完成";
                    pDT = LD.ShowTableInDataGridView_zenjian((ITable)axMapControl1.get_Layer(0), dgvTable, out FieldName);
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public string GetChineseName(string EnglishName)
        {
            switch (EnglishName)
            {
                case "XZQ":
                    return EnglishName + "(行政区)";
                case "GHFW":
                    return EnglishName + "(规划范围)";
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
                case "YJJBNT":
                    return EnglishName + "(永久基本农田)";
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

        private void OpenMDB(string strFullPath)
        {
            // 打开personGeodatabase,并添加图层 
            IWorkspaceFactory pAccessWorkspaceFactory = new AccessWorkspaceFactoryClass();
            // 打开工作空间并遍历数据集 
            IWorkspace pWorkspace = pAccessWorkspaceFactory.OpenFromFile(strFullPath, 0);
            IEnumDataset pEnumDataset = pWorkspace.get_Datasets(ESRI.ArcGIS.Geodatabase.esriDatasetType.esriDTAny);
            pEnumDataset.Reset();
            IDataset pDataset = pEnumDataset.Next();
            int tableflag = 0;
            TreeNode RootNode = new TreeNode();
            while (pDataset != null)
            {
                // 如果数据集是IFeatureDataset,则遍历它下面的子类 
                if (pDataset is IFeatureDataset)
                {
                    pFeatureWorkspace = (IFeatureWorkspace)pAccessWorkspaceFactory.OpenFromFile(strFullPath, 0);
                    pFeatureDataset = pFeatureWorkspace.OpenFeatureDataset(pDataset.Name);
                    IEnumDataset pEnumDataset1 = pFeatureDataset.Subsets;
                    pEnumDataset1.Reset();
                    IDataset pDataset1 = pEnumDataset1.Next();
                    while (pDataset1 != null)
                    {
                        // 如果子类是FeatureClass，则添加到axMapControl1中 
                        if (pDataset1 is IFeatureClass)
                        {
                            pFeatureLayer = new FeatureLayerClass();
                            pFeatureLayer.FeatureClass = pFeatureWorkspace.OpenFeatureClass(pDataset1.Name);
                            pFeatureLayer.Name = pFeatureLayer.FeatureClass.AliasName;
                            axMapControl1.Map.AddLayer(pFeatureLayer);
                            axMapControl1.ActiveView.FocusMap.get_Layer(0).Visible = false;
                            axMapControl1.ActiveView.Refresh();
                            pDataset1 = pEnumDataset1.Next();
                        }
                    }
                    this.uiTab2.SelectedTab = uiTab2.TabPages[0];
                }
                else if (pDataset is IFeatureClass)
                {
                    pFeatureWorkspace = (IFeatureWorkspace)pWorkspace;
                    pFeatureLayer = new FeatureLayerClass();
                    pFeatureLayer.FeatureClass = pFeatureWorkspace.OpenFeatureClass(pDataset.Name);
                    pFeatureLayer.Name = pFeatureLayer.FeatureClass.AliasName;
                    axMapControl1.Map.AddLayer(pFeatureLayer);
                    axMapControl1.ActiveView.FocusMap.get_Layer(0).Visible = true;
                    axMapControl1.ActiveView.Refresh();
                    this.uiTab2.SelectedTab = uiTab2.TabPages[0];
                }
                else
                {
                    if (tableflag == 0)
                    {
                        this.uiTab2.SelectedTab = uiTab2.TabPages[1];
                        RootNode.Text = pWorkspace.PathName;
                        treeView1.Nodes.Add(RootNode);
                    }
                    TreeNode node = new TreeNode();
                    node.Text = pDataset.Name;
                    RootNode.Nodes.Add(node);
                    tableflag = 1;
                }
                pDataset = pEnumDataset.Next();
            }
            treeView1.ExpandAll();
        }

        LoadData LD = new LoadData();
        private void treeView1_DoubleClick(object sender, EventArgs e)
        {
            if (this.treeView1.SelectedNode == null || this.treeView1.SelectedNode.Nodes.Count != 0) return;
            uiStatusBar1.Panels[0].Text = "正在加载数据...";
            // 打开personGeodatabase,并添加图层 
            IWorkspaceFactory pAccessWorkspaceFactory = new AccessWorkspaceFactoryClass();
            // 打开工作空间并遍历数据集 
            IWorkspace pWorkspace = pAccessWorkspaceFactory.OpenFromFile(this.treeView1.SelectedNode.Parent.Text, 0);
            ITable ptable = ((IFeatureWorkspace)pWorkspace).OpenTable(this.treeView1.SelectedNode.Text);
            pDT = LD.ShowTableInDataGridView_zenjian(ptable, dgvTable, out FieldName);
            this.tabMapTableView.SelectedTab = tabMapTableView.TabPages[1];
            uiStatusBar1.Panels[0].Text = "数据加载完成";
        }

        private void btnZoomIn_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            this.axMapControl1.MousePointer = ESRI.ArcGIS.Controls.esriControlsMousePointer.esriPointerZoomIn;
            this._mapAction = MapAction.ZoomIn;
            axToolbarControl1.CurrentTool = null;
        }

        private void btnZoomOut_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            this.axMapControl1.MousePointer = ESRI.ArcGIS.Controls.esriControlsMousePointer.esriPointerZoomOut;
            this._mapAction = MapAction.ZoomOut;
            axToolbarControl1.CurrentTool = null;
        }

        private void btnPan_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            this._mapAction = MapAction.Pan;
            this.axMapControl1.MousePointer = ESRI.ArcGIS.Controls.esriControlsMousePointer.esriPointerHand;
            axToolbarControl1.CurrentTool = null;
        }

        private void btnFullExtent_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            this.axMapControl1.Extent = this.axMapControl1.FullExtent;
            axToolbarControl1.CurrentTool = null;
        }

        private void btnNormal_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            this._mapAction = MapAction.Null;
            this.axMapControl1.MousePointer = ESRI.ArcGIS.Controls.esriControlsMousePointer.esriPointerDefault;
            axToolbarControl1.CurrentTool = null;
        }

        private void axMapControl1_OnMouseDown(object sender, IMapControlEvents2_OnMouseDownEvent e)
        {
            ESRI.ArcGIS.Geometry.IEnvelope _IEnvelope;
            switch (this._mapAction)
            {
                case MapAction.Pan:
                    this.axMapControl1.Pan();
                    break;
                case MapAction.ZoomIn:
                    this.axMapControl1.TrackRectangle();
                    _IEnvelope = this.axMapControl1.Extent;
                    _IEnvelope.Expand(0.5, 0.5, true);
                    this.axMapControl1.Extent = _IEnvelope;
                    break;
                case MapAction.ZoomOut:
                    this.axMapControl1.TrackRectangle();
                    _IEnvelope = this.axMapControl1.Extent;
                    _IEnvelope.Expand(2, 2, true);
                    this.axMapControl1.Extent = _IEnvelope;
                    break;
            }
        }

        private void btnOpenAttributeTable_Click(object sender, EventArgs e)
        {
            if (selectedLayer != null)
            {
                pDT = LD.ShowTableInDataGridView_zenjian((ITable)selectedLayer, dgvTable,out FieldName);
                this.tabMapTableView.SelectedTab = tabMapTableView.TabPages[1];
            }
        }

        private void axTOCControl1_OnMouseDown(object sender, ITOCControlEvents_OnMouseDownEvent e)
        {
            if (e.button == 2)
            {
                ESRI.ArcGIS.Controls.esriTOCControlItem Item = ESRI.ArcGIS.Controls.esriTOCControlItem.esriTOCControlItemNone;
                IBasicMap pBasicMap = null;
                ILayer pLayer = null;
                object other = null;
                object index = null;
                axTOCControl1.HitTest(e.x, e.y, ref Item, ref pBasicMap, ref pLayer, ref other, ref index);          //实现赋值
                selectedLayer = pLayer;
                if (Item == esriTOCControlItem.esriTOCControlItemLayer)           //点击的是图层的话，就显示右键菜单
                {
                    this.contextMenu.Show(axTOCControl1, new System.Drawing.Point(e.x, e.y));
                    //显示右键菜单，并定义其相对控件的位置，正好在鼠标出显示
                }
            }
        }

        private void btnDeleteLayer_Click(object sender, EventArgs e)
        {
            if (selectedLayer != null)
            {
                axMapControl1.Map.DeleteLayer(selectedLayer);
            }
        }


        
        /// <summary>
        /// 按名称获取图层的hook
        /// </summary>
        /// <param name="strLayerName">图层名称</param>
        /// <param name="axMapControl1">MapControl名称</param>
        /// <returns></returns>
        public ILayer GetLayerByName(string strLayerName)
        {
            ILayer pLayer = null;
            for (int i = 0; i <= axMapControl1.LayerCount - 1; i++)
            {
                if (strLayerName == axMapControl1.get_Layer(i).Name)
                {
                    pLayer = axMapControl1.get_Layer(i); break;
                }
            }
            if (pLayer == null)
            {
                string Chinesename = switchName(strLayerName);
                for (int i = 0; i <= axMapControl1.LayerCount - 1; i++)
                {
                    if (Chinesename == axMapControl1.get_Layer(i).Name)
                    {
                        pLayer = axMapControl1.get_Layer(i); break;
                    }
                }

            }
            return pLayer;
        }

        public string switchName(string name)
        {
            switch (name)
            {
                case "CSKFBJ":
                    return "城市开发边界";
                case "HHSM":
                    return "规划河湖水面";
                case "KCDMB":
                    return "行政区划扩充代码表";
                case "QTJSYDQ":
                    return "其他建设用地区";
                case "TDYTQ":
                    return "土地用途区";
                case "JSYDHJBNTGZ ":
                    return "建设用地和基本农田管制";
                case "ZBTZQKB":
                    return "指标调整情况表";
                case "ZBFJB":
                    return "指标分解表";
                case "ZDJSXMYDGHB":
                    return "重点建设项目用地表";
                default:
                    return "";
            }
        }


        private string GetMapUnit(esriUnits _esriMapUnit)
        {
            string sMapUnits = string.Empty;
            switch (_esriMapUnit)
            {
                case esriUnits.esriCentimeters:
                    sMapUnits = "厘米";
                    break;
                case esriUnits.esriDecimalDegrees:
                    sMapUnits = "十进制";
                    break;
                case esriUnits.esriDecimeters:
                    sMapUnits = "分米";
                    break;
                case esriUnits.esriFeet:
                    sMapUnits = "尺";
                    break;
                case esriUnits.esriInches:
                    sMapUnits = "英寸";
                    break;
                case esriUnits.esriKilometers:
                    sMapUnits = "千米";
                    break;
                case esriUnits.esriMeters:
                    sMapUnits = "米";
                    break;
                case esriUnits.esriMiles:
                    sMapUnits = "英里";
                    break;
                case esriUnits.esriMillimeters:
                    sMapUnits = "毫米";
                    break;
                case esriUnits.esriNauticalMiles:
                    sMapUnits = "海里";
                    break;
                case esriUnits.esriPoints:
                    sMapUnits = "点";
                    break;
                case esriUnits.esriUnitsLast:
                    sMapUnits = "UnitsLast";
                    break;
                case esriUnits.esriUnknownUnits:
                    sMapUnits = "米";
                    break;
                case esriUnits.esriYards:
                    sMapUnits = "码";
                    break;
                default:
                    break;
            }
            return sMapUnits;
        }

        private void btnLoadShp_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            OpenFileDialog opfd1 = new OpenFileDialog();
            opfd1.Filter = "shapefile(*.shp)|*.shp|allfile(*.*)|*.*";
            opfd1.Multiselect = false;
            DialogResult diaLres = opfd1.ShowDialog();
            if (diaLres != DialogResult.OK)
                return;
            string path1 = opfd1.FileName;
            //openfiledialog 常规使用
            string pFolder = System.IO.Path.GetDirectoryName(path1);
            string pFileName = System.IO.Path.GetFileName(path1);
            axMapControl1.AddShapeFile(pFolder, pFileName);
        }


        

        private void treeView2_BeforeCollapse(object sender, TreeViewCancelEventArgs e)
        {
            e.Cancel = true;
        }

        





        private void axToolbarControl2_OnMouseUp(object sender, IToolbarControlEvents_OnMouseUpEvent e)
        {
            this._mapAction = MapAction.Null;
            this.axMapControl1.MousePointer = ESRI.ArcGIS.Controls.esriControlsMousePointer.esriPointerDefault;
        }

        private void axToolbarControl1_OnMouseUp(object sender, IToolbarControlEvents_OnMouseUpEvent e)
        {
            this._mapAction = MapAction.Null;
            this.axMapControl1.MousePointer = ESRI.ArcGIS.Controls.esriControlsMousePointer.esriPointerDefault;
        }

        


        /// <summary>
        // 创建要素数据集
        /// </summary>
        /// <param name="workspace"></param>
        /// <param name="code"></param>
        /// <param name="datasetName"></param>
        /// <returns></returns>

        public IFeatureDataset CreateFeatureClass(IWorkspace workspace, IFeatureClass tFeatureClass, string datasetName)
        {
            IFeatureWorkspace featureWorkspace = (IFeatureWorkspace)workspace;
            //创建一个要素集创建一个投影
            ISpatialReferenceFactory spatialRefFactory = new SpatialReferenceEnvironmentClass();
            IDataset dataset = tFeatureClass as IDataset;
            IGeoDataset geoDataset = (IGeoDataset)dataset;
            ISpatialReference spatialReference = geoDataset.SpatialReference;//spatialRefFactory.CreateProjectedCoordinateSystem(code);
            //确定是否支持高精度存储空间
            Boolean supportsHighPrecision = false;
            IWorkspaceProperties workspaceProperties = (IWorkspaceProperties)workspace;
            IWorkspaceProperty workspaceProperty = workspaceProperties.get_Property
                (esriWorkspacePropertyGroupType.esriWorkspacePropertyGroup,
                (int)esriWorkspacePropertyType.esriWorkspacePropSupportsHighPrecisionStorage);
            if (workspaceProperty.IsSupported)
            {
                supportsHighPrecision = Convert.ToBoolean(workspaceProperty.PropertyValue);
            }
            //设置投影精度
            IControlPrecision2 controlPrecision = (IControlPrecision2)spatialReference;
            controlPrecision.IsHighPrecision = supportsHighPrecision;
            //设置容差
            ISpatialReferenceResolution spatialRefResolution = (ISpatialReferenceResolution)spatialReference;
            spatialRefResolution.ConstructFromHorizon();
            spatialRefResolution.SetDefaultXYResolution();
            ISpatialReferenceTolerance spatialRefTolerance = (ISpatialReferenceTolerance)spatialReference;
            spatialRefTolerance.SetDefaultXYTolerance();
            //创建要素集
            IFeatureDataset featureDataset = featureWorkspace.CreateFeatureDataset(datasetName, spatialReference);
            return featureDataset;
        }


        /// <summary>
        /// 获得参照投影的编码
        /// </summary>
        /// <param name="tFeatureLayer"></param>
        /// <returns></returns>
        public int getSpatialReferenceCode(IFeatureClass tFeatureClass)
        {
            IDataset dataset = tFeatureClass as IDataset;
            IGeoDataset geoDataset = (IGeoDataset)dataset;
            int code = geoDataset.SpatialReference.FactoryCode;
            return code;
        }


        /// <summary>
        /// 将Shapefile导入到数据库
        /// </summary>
        /// <param name="pFeaClass"></param>
        /// <param name="pWorkspace"></param>
        /// <param name="tFeatureClass"></param>
        private void importToDB(IFeatureClass pFeaClass, IWorkspace pWorkspace, IFeatureDataset tFeatureClass, string SHPName)
        {
            IFeatureClassDescription featureClassDescription = new FeatureClassDescriptionClass();
            IObjectClassDescription objectClassDescription = featureClassDescription as IObjectClassDescription;
            IFields pFields = pFeaClass.Fields;
            IFieldChecker pFieldChecker = new FieldCheckerClass();
            IEnumFieldError pEnumFieldError = null;
            IFields vFields = null;
            pFieldChecker.ValidateWorkspace = pWorkspace as IWorkspace;
            pFieldChecker.Validate(pFields, out pEnumFieldError, out vFields);
            IFeatureWorkspace featureWorkspace = pWorkspace as IFeatureWorkspace;
            IFeatureClass sdeFeatureClass = null;
            if (sdeFeatureClass == null)
            {
                sdeFeatureClass = tFeatureClass.CreateFeatureClass(SHPName, vFields,
                    objectClassDescription.InstanceCLSID, objectClassDescription.ClassExtensionCLSID,
                    pFeaClass.FeatureType, pFeaClass.ShapeFieldName, "");
                IFeatureCursor featureCursor = pFeaClass.Search(null, true);
                IFeature feature = featureCursor.NextFeature();
                IFeatureCursor sdeFeatureCursor = sdeFeatureClass.Insert(true);
                IFeatureBuffer sdeFeatureBuffer;
                while (feature != null)
                {
                    sdeFeatureBuffer = sdeFeatureClass.CreateFeatureBuffer();
                    IField shpField = new FieldClass();
                    IFields shpFields = feature.Fields;
                    for (int i = 0; i < shpFields.FieldCount; i++)
                    {
                        shpField = shpFields.get_Field(i);
                        if (shpField.Name.Contains("Area") || shpField.Name.Contains("Leng") || shpField.Name.Contains("ID")) continue;
                        int index = sdeFeatureBuffer.Fields.FindField(shpField.Name);
                        if (index != -1)
                        {
                            sdeFeatureBuffer.set_Value(index, feature.get_Value(i));
                        }
                    }
                    sdeFeatureCursor.InsertFeature(sdeFeatureBuffer);
                    sdeFeatureCursor.Flush();
                    feature = featureCursor.NextFeature();
                }
                featureCursor.Flush();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(featureCursor);
            }
        }

        private void btnColorRander_Click(object sender, EventArgs e)
        {
            if (selectedLayer == null)
            {
                return;
            }
            string fieldName = LayerConvertAttr(selectedLayer.Name);
            if (fieldName == "")
            {
                return;
            }
            DefineFeatureColor(selectedLayer, fieldName);
        }

        public void DefineFeatureColor(ILayer player, string fieldName)
        {
            IGeoFeatureLayer m_pGeoFeatureL;
            IUniqueValueRenderer pUniqueValueR;
            IFillSymbol pFillSymbol;
            IColor pNextUniqueColor;
            ITable pTable;
            int lfieldNumber;
            IRow pNextRow;
            IRowBuffer pNextRowBuffer;
            ICursor pCursor;
            IQueryFilter pQueryFilter;
            string codeValue;
            string strNameField = fieldName;
            IMap pMap = this.axMapControl1.Map;
            pMap.ReferenceScale = 0;
            m_pGeoFeatureL = (IGeoFeatureLayer)player;
            pUniqueValueR = new UniqueValueRendererClass();
            pTable = (ITable)m_pGeoFeatureL;
            lfieldNumber = pTable.FindField(strNameField);
            if (lfieldNumber == -1)
            {
                MessageBox.Show("未能找到字段 " + strNameField);
                return;
            }
            //只用一个字段进行单值着色
            pUniqueValueR.FieldCount = 1;
            //用于区分着色的字段
            pUniqueValueR.set_Field(0, strNameField);
            pNextUniqueColor = null;
            //产生查询过滤器对象
            pQueryFilter = new QueryFilterClass();
            pQueryFilter.AddField(strNameField);
            //根据某个字段在表中找出指向所有行的游标对象
            pCursor = pTable.Search(pQueryFilter, true);
            pNextRow = pCursor.NextRow();
            //遍历所有的要素
            while (pNextRow != null)
            {
                pNextRowBuffer = pNextRow;
                //找出Row为“STATES_NAME”的值，即不同的州名
                codeValue = (string)pNextRowBuffer.get_Value(lfieldNumber);
                pNextUniqueColor = getAttrColor(codeValue, fieldName);
                pFillSymbol = new SimpleFillSymbolClass();
                pFillSymbol.Color = pNextUniqueColor;
                //将每次都得的要素字段值和修饰它的符号值放入着色对象中
                pUniqueValueR.AddValue(codeValue, strNameField, (ISymbol)
               pFillSymbol);
                pNextRow = pCursor.NextRow();
            }
            m_pGeoFeatureL.Renderer = (IFeatureRenderer)pUniqueValueR;
            axMapControl1.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeography, null, null);
            pMap.ReferenceScale = 0;
        }

        public string LayerConvertAttr(string Layer)
        {
            switch (Layer)
            {
                case "CSKFBJNGHYT":
                case "城市开发边界内规划用途":
                    return "GHYT";
                case "JSYDKZX":
                case "建设用地控制线":
                    return "LX";
                case "YJJBNT":
                case "永久基本农田":
                    return "YSDM";
                case "JSYDHJBNTGZ2035":
                case "建设用地和基本农田管制":
                    return "GZQLXDM";
                case "JLHDK":
                case "现状建设用地减量化地块":
                    return "SSSX";
                case "JQDLTB":
                case "基期地类图斑":
                    return "DLBM_SX";
                case "STKJKZX":
                case "生态空间控制线":
                    return "GKDJ";
                default:
                    return "";
            }
        }

        public IColor getAttrColor(string AttrValue, string FieldName)
        {
            IColor pcolor = new RgbColorClass();
            if (FieldName == "GHYT")
            {
                switch (AttrValue)
                {
                    case "010":
                        pcolor.RGB = 153 * 65536 + 150 * 256 + 233;
                        return pcolor;
                    case "021":
                        pcolor.RGB = 88 * 65536 + 133 * 256 + 177;
                        return pcolor;
                    case "022":
                        pcolor.RGB = 116 * 65536 + 194 * 256 + 233;
                        return pcolor;
                    case "030":
                        pcolor.RGB = 192 * 65536 + 197 * 256 + 201;
                        return pcolor;
                    case "040":
                        pcolor.RGB = 255 * 65536 + 239 * 256 + 150;
                        return pcolor;
                    default:
                        pcolor.RGB = 0 * 65536 + 0 * 256 + 0;
                        return pcolor;
                }
            }
            else if (FieldName == "LX")
            {
                switch (AttrValue)
                {
                    case "其他建设用地区":
                        pcolor.RGB = 153 * 65536 + 150 * 256 + 233;
                        return pcolor;
                    case "城市开发边界内建设用地":
                        pcolor.RGB = 219 * 65536 + 89 * 256 + 245;
                        return pcolor;
                    default:
                        pcolor.RGB = 0 * 65536 + 0 * 256 + 0;
                        return pcolor;
                }
            }
            else if (FieldName == "YSDM")
            {
                switch (AttrValue)
                {
                    default:
                        pcolor.RGB = 0 * 65536 + 232 * 256 + 245;
                        return pcolor;
                }
            }
            else if (FieldName == "GZQLXDM")
            {
                switch (AttrValue)
                {
                    case "011":
                        pcolor.RGB = 153 * 65536 + 150 * 256 + 233;
                        return pcolor;
                    case "012":
                        pcolor.RGB = 154 * 65536 + 89 * 256 + 167;
                        return pcolor;
                    case "031":
                    case "032":
                    case "033":
                        pcolor.RGB = 146 * 65536 + 238 * 256 + 244;
                        return pcolor;
                    case "040":
                        pcolor.RGB = 24 * 65536 + 111 * 256 + 46;
                        return pcolor;
                    default:
                        pcolor.RGB = 0 * 65536 + 0 * 256 + 0;
                        return pcolor;
                }
            }
            else if (FieldName == "SSSX")
            {
                switch (AttrValue)
                {
                    case "近期":
                        pcolor.RGB = 0 * 65536 + 152 * 256 + 230;
                        return pcolor;
                    case "远期":
                        pcolor.RGB = 116 * 65536 + 194 * 256 + 233;
                        return pcolor;
                    default:
                        pcolor.RGB = 0 * 65536 + 0 * 256 + 0;
                        return pcolor;
                }
            }
            else if (FieldName == "DLBM_SX")
            {
                Regex re = new Regex(@"11(\w+)"); //以11开头的单词
                if (re.IsMatch(AttrValue) || AttrValue.Contains("K"))
                {
                    pcolor.RGB = 100 * 65536 + 255 * 256 + 255;
                    return pcolor;
                }
                re = new Regex(@"12(\w+)");
                if (re.IsMatch(AttrValue))
                {
                    pcolor.RGB = 190 * 65536 + 255 * 256 + 255;
                    return pcolor;
                }
                re = new Regex(@"13(\w+)");
                if (re.IsMatch(AttrValue))
                {
                    pcolor.RGB = 115 * 65536 + 255 * 256 + 164;
                    return pcolor;
                }
                if (AttrValue == "155")
                {
                    pcolor.RGB = 242 * 65536 + 219 * 256 + 197;
                    return pcolor;
                }
                if (AttrValue == "154")
                {
                    pcolor.RGB = 223 * 65536 + 217 * 256 + 204;
                    return pcolor;
                }
                if (AttrValue == "151" || AttrValue == "152" || AttrValue == "157" || AttrValue == "158")
                {
                    pcolor.RGB = 126 * 65536 + 255 * 256 + 209;
                    return pcolor;
                }
                re = new Regex(@"14(\w+)");
                if (AttrValue == "153" || AttrValue == "156" || re.IsMatch(AttrValue))
                {
                    pcolor.RGB = 0 * 65536 + 170 * 256 + 112;
                    return pcolor;
                }
                if (AttrValue == "251" || AttrValue == "252")
                {
                    pcolor.RGB = 144 * 65536 + 139 * 256 + 248;
                    return pcolor;
                }
                re = new Regex(@"22(\w+)");
                if (re.IsMatch(AttrValue))
                {
                    pcolor.RGB = 135 * 65536 + 146 * 256 + 208;
                    return pcolor;
                }
                if (AttrValue == "253" || AttrValue == "254")
                {
                    pcolor.RGB = 200 * 65536 + 171 * 256 + 255;
                    return pcolor;
                }
                re = new Regex(@"26(\w+)");
                if (re.IsMatch(AttrValue))
                {
                    pcolor.RGB = 178 * 65536 + 178 * 256 + 178;
                    return pcolor;
                }
                re = new Regex(@"21(\w+)");
                if (re.IsMatch(AttrValue))
                {
                    pcolor.RGB = 0 * 65536 + 0 * 256 + 244;
                    return pcolor;
                }
                re = new Regex(@"23(\w+)");
                if (re.IsMatch(AttrValue))
                {
                    pcolor.RGB = 255 * 65536 + 1 * 256 + 255;
                    return pcolor;
                }
                re = new Regex(@"24(\w+)");
                if (re.IsMatch(AttrValue))
                {
                    pcolor.RGB = 146 * 65536 + 55 * 256 + 181;
                    return pcolor;
                }
                re = new Regex(@"27(\w+)");
                if (re.IsMatch(AttrValue))
                {
                    pcolor.RGB = 200 * 65536 + 101 * 256 + 241;
                    return pcolor;
                }
                re = new Regex(@"28(\w+)");
                if (re.IsMatch(AttrValue))
                {
                    pcolor.RGB = 171 * 65536 + 153 * 256 + 192;
                    return pcolor;
                }
                if (AttrValue == "321" || AttrValue == "322")
                {
                    pcolor.RGB = 255 * 65536 + 239 * 256 + 150;
                    return pcolor;
                }
                if (AttrValue == "323" || AttrValue == "324")
                {
                    pcolor.RGB = 243 * 65536 + 220 * 256 + 152;
                    return pcolor;
                }
                re = new Regex(@"31(\w+)");
                Regex re2 = new Regex(@"33(\w+)");
                if (re.IsMatch(AttrValue) || re2.IsMatch(AttrValue))
                {
                    pcolor.RGB = 225 * 65536 + 225 * 256 + 225;
                    return pcolor;
                }
            }
            else if (FieldName == "GKDJ")
            {
                switch (AttrValue)
                {
                    case "01":
                        pcolor.RGB = 104 * 65536 + 129 * 256 + 43;
                        return pcolor;
                    case "02":
                        pcolor.RGB = 136 * 65536 + 177 * 256 + 55;
                        return pcolor;
                    case "03":
                        pcolor.RGB = 200 * 65536 + 241 * 256 + 213;
                        return pcolor;
                    case "04":
                        pcolor.RGB = 64 * 65536 + 166 * 256 + 85;
                        return pcolor;
                    default:
                        pcolor.RGB = 0 * 65536 + 0 * 256 + 0;
                        return pcolor;
                }
            }
            pcolor.RGB = 0 * 65536 + 0 * 256 + 0;
            return pcolor;

        }

        
        private IFeatureClass AddFeatureToFeatureClass(IFeatureClass pFeatureClass, IFeature pFeature)
        {
            IFeatureCursor pFeatureCursor = pFeatureClass.Insert(true);
            IFeatureBuffer pFeatureBuffer = pFeatureClass.CreateFeatureBuffer();
            IFields pFields = pFeatureClass.Fields;
            for (int i = 1; i <= pFeature.Fields.FieldCount - 1; i++)
            {
                IField pField = pFields.get_Field(i);
                if (pField.Type == esriFieldType.esriFieldTypeGeometry)
                {
                    //pFeatureBuffer.set_Value(i, Convert.ToString(pFeature.get_Value(i)));
                    pFeatureBuffer.set_Value(i, pFeature.ShapeCopy);
                }
                else
                {
                    switch (pField.Type)
                    {
                        case esriFieldType.esriFieldTypeInteger:
                            pFeatureBuffer.set_Value(i, Convert.ToInt32(pFeature.get_Value(i)));

                            break;
                        case esriFieldType.esriFieldTypeDouble:
                            //pFeatureBuffer.set_Value(i, Convert.ToDouble(pFeature.get_Value(i)));
                            break;
                        case esriFieldType.esriFieldTypeString:
                            pFeatureBuffer.set_Value(i, Convert.ToString(pFeature.get_Value(i)));
                            break;
                        default:
                            break;
                    }
                }
            }
            pFeatureCursor.InsertFeature(pFeatureBuffer);
            return pFeatureClass;
        }

        public void addFields(IFeature pFeature, IFeatureClass pFeatureClass, int zdCount)  //传入待修改要素类及样本要素
        {
            IFields ppFileds = pFeatureClass.Fields;
            IFieldsEdit ppFieldsEdit = (IFieldsEdit)ppFileds;
            IFields pFields = pFeature.Fields;
            while (pFields.FieldCount > pFeatureClass.Fields.FieldCount)
            {

                for (int i = zdCount; i < pFeature.Fields.FieldCount; i++)
                {

                    IField ppField = new FieldClass();
                    IField pField = pFields.get_Field(i);
                    IFieldEdit pFieldEdit = (IFieldEdit)ppField;
                    if (i == 0)
                    {
                        //pFeatureClass.GetFeature(1).set_Value(1, pFields.get_Field(1));
                        pFieldEdit.Name_2 = pFields.get_Field(i).Name;
                        pFeatureClass.GetFeature(i).Delete();
                        pFeatureClass.GetFeature(i).Store();

                    }
                    else
                    {
                        pFieldEdit.Name_2 = pFields.get_Field(i).Name;
                        pFieldEdit.Type_2 = pField.Type;
                        pFieldEdit.Editable_2 = true;
                        ppFieldsEdit.AddField(ppField);
                    }
                }

            }
        }

        private void buttonCommand1_Click_1(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            SearchForm _SelectbyAttributeFrm = new SearchForm();
            _SelectbyAttributeFrm.SqlOK += _SelectbyAttributeFrm_SqlOK;
            List<ILayer> _layerInfo = new List<ILayer>();
            for (int layerIndex = 0; layerIndex < this.axMapControl1.LayerCount; layerIndex++)
            {
                _layerInfo.Add(this.axMapControl1.get_Layer(layerIndex));
            }

            _SelectbyAttributeFrm.ShowInfo(_layerInfo);
        }

        private IRgbColor getRGB(int r, int g, int b)
        {
            IRgbColor pRgbColor;
            pRgbColor = new RgbColorClass();
            pRgbColor.Red = r;
            pRgbColor.Green = g;
            pRgbColor.Blue = b;
            return pRgbColor;
        }

        private void _SelectbyAttributeFrm_SqlOK(object sender, SQLFileterEventArgs e)
        {
            IFeatureSelection layer = this.axMapControl1.get_Layer(e.LayerIndex) as IFeatureSelection;
            IQueryFilter pQueryFilter = new QueryFilterClass();
            pQueryFilter.WhereClause = e.SQL;//过滤条件，查询表达式
            layer.SelectFeatures(pQueryFilter, esriSelectionResultEnum.esriSelectionResultNew, false);

            ISimpleFillSymbol SFS = new SimpleFillSymbolClass();
            ISimpleLineSymbol ILS = new SimpleLineSymbolClass();
            SFS.Style = esriSimpleFillStyle.esriSFSSolid;
            SFS.Color = getRGB(255, 0, 0);
            ILS.Color = getRGB(0, 255, 0);
            ILS.Style = esriSimpleLineStyle.esriSLSSolid;
            ILS.Width = 13;
            SFS.Outline = ILS;
            layer.SelectionSymbol = SFS as ISymbol;
            //this.dgvTable.Columns;
            this.axMapControl1.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphicSelection, null, null);
            this.axMapControl1.Refresh();

            // 这个是你查询出来的DataTable中的行集合
            DataRow[] rowsinDataTable = pDT.Select(e.SQL_2);
            dgvTable.MultiSelect = true;
            dgvTable.ClearSelection();
            foreach (DataRow r in rowsinDataTable)
            {
                foreach (DataGridViewRow row in dgvTable.Rows)
                {
                    // 假设ID为第一个单元格,比较他们之间的值
                    if (r["OBJECTID"] == row.Cells[0].Value)
                    {
                        // 相等就代表你查询出的数据行在DataGridView 中存在，并选中对应的数据行
                        row.Selected = true;
                    }

                }
            }

            DataView dv = new DataView((DataTable)dgvTable.DataSource);
            dv.RowFilter = e.SQL_2;
            pDTSearch = dv.ToTable();
            dgvSearch.DataSource = pDTSearch;
            for (int i = 0; i < FieldName.Count; i++)
            {
                dgvSearch.Columns[i].HeaderText = FieldName[i];
            }
            //DateTime dt;
            //DateTimeFormatInfo dtFormat = new DateTimeFormatInfo();
            //dtFormat.ShortDatePattern = "yyyy/MM/dd";
            //dt = Convert.ToDateTime("2011/05/26", dtFormat);
            foreach (DataGridViewColumn column in dgvSearch.Columns)
            { column.SortMode = DataGridViewColumnSortMode.NotSortable; }
        }

        static DataTable pDTSearch;

        private void dgvTable_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int OID_index = -1;
            for (int count = 0; count < dgvTable.Columns.Count; count++)
            {
                if (dgvTable.Columns[count].HeaderText == "OBJECTID") { OID_index = count; break; }
            }
            if (OID_index == -1) { return; };
            ILayer player = axMapControl1.get_Layer(0);
            int OID = int.Parse(dgvTable.SelectedRows[0].Cells[OID_index].Value.ToString());
            IArea pArea = (player as IFeatureLayer).FeatureClass.GetFeature(OID).Shape as IArea;
            IPoint iPnt = pArea.LabelPoint;
            axMapControl1.Extent = (player as IFeatureLayer).FeatureClass.GetFeature(OID).Shape.Envelope;
            axMapControl1.CenterAt(iPnt);
            axMapControl1.Refresh();
        }

        private void dgvSearch_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dgvTable.SelectedRows.Count == 0) { return; }
            int OID_index = -1;
            for (int count = 0; count < dgvTable.Columns.Count; count++)
            {
                if (dgvTable.Columns[count].HeaderText == "OBJECTID") { OID_index = count; break; }
            }
            if (OID_index == -1) { return; };
            ILayer player = axMapControl1.get_Layer(0);
            int OID = int.Parse(dgvSearch.SelectedRows[0].Cells[OID_index].Value.ToString());
            IArea pArea = (player as IFeatureLayer).FeatureClass.GetFeature(OID).Shape as IArea;
            IPoint iPnt = pArea.LabelPoint;
            axMapControl1.Extent = (player as IFeatureLayer).FeatureClass.GetFeature(OID).Shape.Envelope;
            axMapControl1.CenterAt(iPnt);
            axMapControl1.Refresh();
        }

        private void dgvSearch_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            foreach (DataGridViewColumn column in dgvSearch.Columns)
            { column.SortMode = DataGridViewColumnSortMode.NotSortable; }
            Change _frmChange = new Change(FieldName[e.ColumnIndex]);
            dgvSearch_select = e.ColumnIndex;
            _frmChange.ChangeOK += _frmChange_ChangeOK;
            _frmChange.Show();
        }

        Revise revise = new Revise();
        int dgvSearch_select;
        private void _frmChange_ChangeOK(object sender, ChangeEventArgs e)
        {
            for (int i = 0; i < dgvSearch.Rows.Count; i++)
            {
                dgvSearch[dgvSearch_select,i].Value = e.field_value;
            }
        }

        private void dgvTable_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //if (e.Button == MouseButtons.Right && cbcActivateAlter.Checked == true)
            //{
            //    this.contextMenuStrip1.Show(splitContainer4.Panel1, new System.Drawing.Point(e.Location.X, e.Location.Y));
            //    //显示右键菜单，并定义其相对控件的位置，正好在鼠标出显示
            //    dgvTable_select_Index = e.ColumnIndex;
            //}
        }
        int dgvTable_select_Index;
        private void AddToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Addname addname = new Addname(tempFeatureLayer as IFeatureLayer,this,FieldName);
            addname.Show();
        }

        IFeatureLayer tempFeatureLayer;
        private void cbcActivateAlter_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            if (axMapControl1.Map.LayerCount == 0) { return; }
            if (cbcActivateAlter.Checked == true)
            {
                this.btnadd.Enabled = false;
                this.dgvSearch.ReadOnly = false;
                string fileNameExt = DateTime.Now.ToString("yyyyMMddHHmmss") + ".mdb";
                string filePath = System.IO.Directory.GetCurrentDirectory();
                IWorkspaceFactory pWorksapceFactory = new AccessWorkspaceFactory();
                IWorkspaceName worksapcename = pWorksapceFactory.Create(filePath, fileNameExt, null, 0);
                IName name = worksapcename as IName;
                IWorkspace pWorkspace = name.Open() as IWorkspace;
                IFeatureLayer mCphFeatureLayer = axMapControl1.get_Layer(0) as IFeatureLayer;//这是获得要入库的shapefile，获取其FeatureLayer即可
                //2.创建要素数据集
                IFeatureClass pCphFeatureClass = mCphFeatureLayer.FeatureClass;
                //int code = getSpatialReferenceCode(pCphFeatureClass);//参照投影的代号
                string datasetName = pCphFeatureClass.AliasName;//要素数据集的名称
                IFeatureDataset pCphDataset = CreateFeatureClass(pWorkspace, pCphFeatureClass, datasetName);
                //3.导入SHP到要素数据集(
                importToDB(pCphFeatureClass, pWorkspace, pCphDataset, pCphFeatureClass.AliasName);

                // 打开personGeodatabase,并添加图层 
                IWorkspaceFactory pAccessWorkspaceFactory = new AccessWorkspaceFactoryClass();
                // 打开工作空间并遍历数据集 
                IWorkspace temp_Workspace = pAccessWorkspaceFactory.OpenFromFile(filePath + "/" + fileNameExt, 0);
                IEnumDataset pEnumDataset = pWorkspace.get_Datasets(ESRI.ArcGIS.Geodatabase.esriDatasetType.esriDTAny);
                pEnumDataset.Reset();
                IDataset pDataset = pEnumDataset.Next();

                if (pDataset is IFeatureDataset)
                {
                    pFeatureWorkspace = (IFeatureWorkspace)pAccessWorkspaceFactory.OpenFromFile(filePath + "/" + fileNameExt, 0);
                    pFeatureDataset = pFeatureWorkspace.OpenFeatureDataset(pDataset.Name);
                    IEnumDataset pEnumDataset1 = pFeatureDataset.Subsets;
                    pEnumDataset1.Reset();
                    IDataset pDataset1 = pEnumDataset1.Next();
                    if (pDataset1 is IFeatureClass)
                    {
                        tempFeatureLayer = new FeatureLayerClass();
                        tempFeatureLayer.FeatureClass = pFeatureWorkspace.OpenFeatureClass(pDataset1.Name);
                        tempFeatureLayer.Name = pFeatureLayer.FeatureClass.AliasName;
                    }
                }
            }
            else
            {
                this.btnadd.Enabled = false;
                this.dgvSearch.ReadOnly = true;
            }
            
        }

        public void UpdateFTOnDV(ILayer player, DataTable pdatatable, int[] array)
        {
            IFeatureLayer pFTClass = player as IFeatureLayer;
            ITable pTable = pFTClass as ITable;
            ICursor pCursor;
            IRow pRow;
            pCursor = pTable.GetRows(array, false);
            for (int i = 0; i < array.Length; i++)
            {
                pRow = pCursor.NextRow();
                int k = array[i];
                for (int j = 2; j < pdatatable.Columns.Count; j++)
                {
                    object pgridview = pdatatable.Rows[k][j];
                    object prow = pRow.get_Value(j);
                    if (prow.ToString() != pgridview.ToString())
                    {
                        pRow.set_Value(j, pgridview);
                        pRow.Store();
                    }
                }

            }

            MessageBox.Show("数据保存成功！");
        }

        private void btnSave_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            dgvSearch.CurrentCell = null;
            for (int i=0; i < this.dgvSearch.RowCount; i++)
            {
                string oid = dgvSearch.Rows[i].Cells[0].Value.ToString();

                IFeatureLayer pfeaturelayer = axMapControl1.get_Layer(0) as IFeatureLayer;
                
                //找到要素
                IQueryFilter pQueryFilter = new QueryFilter();
                pQueryFilter.WhereClause = "OBJECTID = " + oid;

                IFeatureCursor pFeatureCur = pfeaturelayer.Search(pQueryFilter, false);

                IFeature pFeature = null;

                pFeature = pFeatureCur.NextFeature();

                if (null == pFeature){}
                else
                {
                    IFields pFields = pFeature.Fields;
                    IFeatureClass pFeatureClass = pfeaturelayer.FeatureClass;
                    for (int j = 0; j < pFeature.Fields.FieldCount; j++)
                    {
                        if (pFeature.Fields.get_Field(j).Type != esriFieldType.esriFieldTypeString) { continue; }
                        pFeature.set_Value(j, dgvSearch[j,i].Value);
                    }
                    pFeature.Store();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(pFeatureCur);
                }
            }
            pDT = LD.ShowTableInDataGridView_zenjian(axMapControl1.get_Layer(0) as ITable, dgvTable, out FieldName);
        }
        //导出附表1


        private void btnPreview1_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            pDT.DefaultView.Sort = "区名 desc";
            DataTable pDataTable1 = new DataTable();//建立一个table
            string[] FieldName = new string[] { "Index", "District", "RegionName", "RegionIndex", "AdmitedTime", "ImplementSpan", "ProcessState" };
            //表格2 string[] FieldName = new string[] { "Index2", "RegionName2", "RegionIndex2", "SetupRegion_new", "Area_new", "OutRegion_new", "OutArea","OutTarget","OldRegion_old","OldArea","InvolveFarmers","ApprovalTime","ApprovalIndex","Deadline","BuildingArea","Plough" };

            //表格3 string[] FieldName = new string[] { "RegionName3", "RegionIndex3", "InvolvedTown", "PlanuseBuilding_all", "PlanusePlough_all", "RealuseBuilding_all", "RealusePlough_all", "PlanuseBuilding_setup", "PlanusePlough_setup", "RealuseBuilding_setup", "RealusePlough_setup", "PlanuseBuilding_out", "PlanusePlough_out", "RealuseBuilding_out", "RealusePlough_out","Planreturnbuildingarea","Planreturnplough","Realreturnbuildingarea","Realreturnplough" };
            //string[] FieldName = new string[] { "Index", "District", "RegionName", "RegionIndex", "AdmitedTime", "ImplementSpan", "ProcessState" };
            //string[] FieldName = new string[] { "Index", "District", "RegionName", "RegionIndex", "AdmitedTime", "ImplementSpan", "ProcessState" };
            for (int i = 0; i < FieldName.Length; i++)
            {
                pDataTable1.Columns.Add(FieldName[i]);
            }
            //pDataTable1.Columns.Add("count");
            DataTable dtName = pDT.DefaultView.ToTable(true, "XMMC");
            for (int i = 0; i < dtName.Rows.Count; i++)
            {
                DataRow[] rows = pDT.Select("XMMC='" + dtName.Rows[i][0] + "'");
                //temp用来存储筛选出来的数据
                //DataTable temp = pDataTable1.Clone();
                //foreach (DataRow row in rows)
                //{
                //    temp.Rows.Add(row.ItemArray);
                //}

                DataRow StrRow = pDataTable1.NewRow();
                StrRow[0] = (i + 1).ToString();
                StrRow[1] = rows[0]["区名"].ToString();
                StrRow[2] = rows[0]["XMMC"].ToString();
                StrRow[3] = rows[0]["XMBH"].ToString();
                StrRow[4] = rows[0]["项目批复时间"].ToString();
                StrRow[5] = rows[0]["实施期限"].ToString();
                StrRow[6] = rows[0]["项目进度概况"].ToString();
                //StrRow[7] = rows.Count().ToString();
                pDataTable1.Rows.Add(StrRow);
            } 


            //for (int i = 0; i < pDT.Rows.Count; i++)
            //{
            //    DataRow pRow = pDataTable1.NewRow();
            //    string[] StrRow = new string[7];
            //    StrRow[0] = (i + 1).ToString();
            //    StrRow[1] = pDT.Rows[i]["区名"].ToString();
            //    StrRow[2] = pDT.Rows[i]["XMMC"].ToString();
            //    StrRow[3] = pDT.Rows[i]["XMBH"].ToString();
            //    StrRow[4] = pDT.Rows[i]["项目批复时间"].ToString();
            //    StrRow[5] = pDT.Rows[i]["实施期限"].ToString();
            //    StrRow[6] = pDT.Rows[i]["项目进度概况"].ToString();
            //    pRow.ItemArray = StrRow;
            //    pDataTable1.Rows.Add(pRow);
            //}
            dgv_Table1.DataSource = pDataTable1;
        }

        public double calcu_sum(DataRow[] rows)
        {
            double sum = 0;
            for(int i = 0;i<rows.Count();i++)
            {
                double result = 0;
                if (Double.TryParse(rows[i]["Shape_Area"].ToString(),out result)) 
                {
                    sum += result;
                }
            }
            return sum;
        }
        private void btnPreview2_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            pDT.DefaultView.Sort = "区名 desc";
            DataTable pDataTable1 = new DataTable();//建立一个table
            string[] FieldName = new string[] { "Index2", "RegionName2", "RegionIndex2", "SetupRegion_new", "Area_new", "OutRegion_new", "OutArea","OutTarget","OldRegion_old","OldArea","InvolveFarmers","ApprovalTime","ApprovalIndex","Deadline","BuildingArea","Plough" };

            //表格3 string[] FieldName = new string[] { "RegionName3", "RegionIndex3", "InvolvedTown", "PlanuseBuilding_all", "PlanusePlough_all", "RealuseBuilding_all", "RealusePlough_all", "PlanuseBuilding_setup", "PlanusePlough_setup", "RealuseBuilding_setup", "RealusePlough_setup", "PlanuseBuilding_out", "PlanusePlough_out", "RealuseBuilding_out", "RealusePlough_out","Planreturnbuildingarea","Planreturnplough","Realreturnbuildingarea","Realreturnplough" };
            //string[] FieldName = new string[] { "Index", "District", "RegionName", "RegionIndex", "AdmitedTime", "ImplementSpan", "ProcessState" };
            //string[] FieldName = new string[] { "Index", "District", "RegionName", "RegionIndex", "AdmitedTime", "ImplementSpan", "ProcessState" };

            for (int i = 0; i < FieldName.Length; i++)
            {
                pDataTable1.Columns.Add(FieldName[i]);
            }
            DataTable dtResult = pDT.Clone();
            DataTable dtName = pDT.DefaultView.ToTable(true, "XMMC");
            for (int i = 0; i < dtName.Rows.Count; i++)
            {
                DataRow[] rows = pDT.Select("XMMC='" + dtName.Rows[i][0] + "'");

                //temp用来存储筛选出来的数据
                DataTable temp = dtResult.Clone();
                foreach (DataRow row in rows)
                {
                    temp.Rows.Add(row.ItemArray);
                }
                DataRow[] anzhi_rows = temp.Select("DKLX='安置地块'");
                DataRow[] churang_rows = temp.Select("DKLX='出让地块'");
                DataRow[] chaijiu_rows = temp.Select("DKLX='拆旧地块'");

                DataRow pRow = pDataTable1.NewRow();
                string[] StrRow = new string[16];
                StrRow[0] = (i + 1).ToString();
                StrRow[1] = rows[0]["XMMC"].ToString();
                StrRow[2] = rows[0]["XMBH"].ToString();
                StrRow[3] = anzhi_rows.Count().ToString();
                StrRow[4] = calcu_sum(anzhi_rows).ToString();
                StrRow[5] = churang_rows.Count().ToString();
                StrRow[6] = calcu_sum(churang_rows).ToString();
                StrRow[8] = chaijiu_rows.Count().ToString();
                StrRow[9] = calcu_sum(chaijiu_rows).ToString();
                StrRow[11] = rows[0]["批复下达时间"].ToString();
                StrRow[12] = rows[0]["批复文号"].ToString();
                StrRow[13] = rows[0]["实施期限"].ToString();
                StrRow[14] = rows[0]["批复文号"].ToString();
                pRow.ItemArray = StrRow;
                pDataTable1.Rows.Add(pRow);
            }
            dgv_Table2.DataSource = pDataTable1;
        }

        public void export(Janus.Windows.Ribbon.ButtonCommand btn)
        {
            DataGridView datagrid;
            switch (btn.Name)
            {
                case "btnExport1":
                    datagrid = dgv_Table1;
                    break;
                case "btnExport2":
                    datagrid = dgv_Table2;
                    break;
                case "btnExport3":
                    datagrid = dgv_Table3;
                    break;
                case "btnExport4":
                    datagrid = dgv_Table4;
                    break;
                default:
                    datagrid = dgv_Table5;
                    break;
            }

            if (datagrid.IsCurrentCellInEditMode == true)
            {
                datagrid.CurrentCell = null;
            }
            string filePath = "";
            SaveFileDialog s = new SaveFileDialog();
            s.Title = "保存Excel文件";
            s.Filter = "Excel文件(*.xlsx)|*.xlsx";
            s.FilterIndex = 1;
            if (s.ShowDialog() == DialogResult.OK)
            {
                filePath = s.FileName;

                DataTable tmpErrorDataTable = new DataTable("ErrorDT");
                DataTable modelTable = new DataTable("ModelTable");
                for (int column = 0; column < datagrid.Columns.Count; column++)
                {
                    if (datagrid.Columns[column].Visible == true)
                    {
                        DataColumn tempColumn = new DataColumn(datagrid.Columns[column].HeaderText, typeof(string));
                        tmpErrorDataTable.Columns.Add(tempColumn);

                        DataColumn modelColumn = new DataColumn(datagrid.Columns[column].Name, typeof(string));
                        modelTable.Columns.Add(modelColumn);
                    }
                }
                for (int row = 0; row < datagrid.Rows.Count; row++)
                {

                    DataRow tempRow = tmpErrorDataTable.NewRow();
                    for (int i = 0; i < tmpErrorDataTable.Columns.Count; i++)
                    {
                        tempRow[i] = datagrid.Rows[row].Cells[modelTable.Columns[i].ColumnName].Value;
                    }
                    tmpErrorDataTable.Rows.Add(tempRow);
                }
                if (tmpErrorDataTable == null)
                {
                    return;
                }
                //第二步：导出dataTable到Excel  
                long rowNum = tmpErrorDataTable.Rows.Count;//行数  
                int columnNum = tmpErrorDataTable.Columns.Count;//列数  
                Excel.Application m_xlApp = new Excel.Application();
                m_xlApp.DisplayAlerts = false;//不显示更改提示  
                m_xlApp.Visible = false;
                Excel.Workbooks workbooks = m_xlApp.Workbooks;
                Excel.Workbook workbook = workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];//取得sheet1  
                try
                {
                    string[,] datas = new string[rowNum + 1, columnNum];
                    for (int i = 0; i < columnNum; i++) //写入字段  
                        datas[0, i] = tmpErrorDataTable.Columns[i].Caption;
                    //Excel.Range range = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[1, columnNum]);  
                    Excel.Range range = m_xlApp.Range[worksheet.Cells[1, 1], worksheet.Cells[1, columnNum]];
                    range.Interior.ColorIndex = 15;//15代表灰色  
                    range.Font.Bold = true;
                    range.Font.Size = 10;
                    int r = 0;
                    for (r = 0; r < rowNum; r++)
                    {
                        for (int i = 0; i < columnNum; i++)
                        {
                            object obj = tmpErrorDataTable.Rows[r][tmpErrorDataTable.Columns[i].ToString()];
                            datas[r + 1, i] = obj == null ? "" : "'" + obj.ToString().Trim();//在obj.ToString()前加单引号是为了防止自动转化格式  
                        }
                        System.Windows.Forms.Application.DoEvents();
                        //添加进度条  
                    }
                    //Excel.Range fchR = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[rowNum + 1, columnNum]);  
                    Excel.Range fchR = m_xlApp.Range[worksheet.Cells[1, 1], worksheet.Cells[rowNum + 1, columnNum]];
                    fchR.Value2 = datas;
                    worksheet.Columns.EntireColumn.AutoFit();//列宽自适应。  
                    //worksheet.Name = "dd";  
                    //m_xlApp.WindowState = Excel.XlWindowState.xlMaximized;
                    m_xlApp.Visible = false;
                    // = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[rowNum + 1, columnNum]);  
                    range = m_xlApp.Range[worksheet.Cells[1, 1], worksheet.Cells[rowNum + 1, columnNum]];
                    //range.Interior.ColorIndex = 15;//15代表灰色  
                    range.Font.Size = 9;
                    range.RowHeight = 14.25;
                    range.Borders.LineStyle = 1;
                    range.HorizontalAlignment = 1;
                    workbook.Saved = true;
                    workbook.SaveCopyAs(filePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("导出异常：" + ex.Message, "导出异常", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    m_xlApp.Workbooks.Close();
                    m_xlApp.Workbooks.Application.Quit();
                    m_xlApp.Application.Quit();
                    m_xlApp.Quit();
                    return;
                }
                finally
                {
                    //EndReport();
                }
                m_xlApp.Workbooks.Close();
                m_xlApp.Workbooks.Application.Quit();
                m_xlApp.Application.Quit();
                m_xlApp.Quit();
                MessageBox.Show("导出成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\n导出成功，路径为" + filePath + "\r\n" });
            }
            else { return; }
        }

        private void btnPreview3_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            pDT.DefaultView.Sort = "区名 desc";
            DataTable pDataTable1 = new DataTable();//建立一个table
            string[] FieldName = new string[] { "RegionName3", "RegionIndex3", "InvolvedTown", "PlanuseBuilding_all", "PlanusePlough_all", "RealuseBuilding_all", "RealusePlough_all", "PlanuseBuilding_setup", "PlanusePlough_setup", "RealuseBuilding_setup", "RealusePlough_setup", "PlanuseBuilding_out", "PlanusePlough_out", "RealuseBuilding_out", "RealusePlough_out", "Planreturnbuildingarea", "Planreturnplough", "Realreturnbuildingarea", "Realreturnplough" };

            //表格3 string[] FieldName = new string[] { "RegionName3", "RegionIndex3", "InvolvedTown", "PlanuseBuilding_all", "PlanusePlough_all", "RealuseBuilding_all", "RealusePlough_all", "PlanuseBuilding_setup", "PlanusePlough_setup", "RealuseBuilding_setup", "RealusePlough_setup", "PlanuseBuilding_out", "PlanusePlough_out", "RealuseBuilding_out", "RealusePlough_out","Planreturnbuildingarea","Planreturnplough","Realreturnbuildingarea","Realreturnplough" };
            //string[] FieldName = new string[] { "Index", "District", "RegionName", "RegionIndex", "AdmitedTime", "ImplementSpan", "ProcessState" };
            //string[] FieldName = new string[] { "Index", "District", "RegionName", "RegionIndex", "AdmitedTime", "ImplementSpan", "ProcessState" };

            for (int i = 0; i < FieldName.Length; i++)
            {
                pDataTable1.Columns.Add(FieldName[i]);
            }
            DataTable dtResult = pDT.Clone();
            DataTable dtName = pDT.DefaultView.ToTable(true, "XMBH");
            for (int i = 0; i < dtName.Rows.Count; i++)
            {
                DataRow[] rows = pDT.Select("XMBH='" + dtName.Rows[i][0] + "'");

                //temp用来存储筛选出来的数据
                DataTable temp = dtResult.Clone();
                foreach (DataRow row in rows)
                {
                    temp.Rows.Add(row.ItemArray);
                }

                DataRow pRow = pDataTable1.NewRow();
                string[] StrRow = new string[19];
                StrRow[0] = rows[0]["区名"].ToString();
                StrRow[1] = rows[0]["XMBH"].ToString();

                StrRow[15] = rows[0]["指标归还情况_归还建设用地面积"].ToString();
                StrRow[16] = rows[0]["指标归还情况_归还耕地面积"].ToString();
                pRow.ItemArray = StrRow;
                pDataTable1.Rows.Add(pRow);
            }
            dgv_Table3.DataSource = pDataTable1;
        }
        private void btnPreview4_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            pDT.DefaultView.Sort = "区名 desc";
            DataTable pDataTable1 = new DataTable();//建立一个table
            string[] FieldName = new string[] { "RegionName4", "ProjectregionName", "Counts", "Area_new4", "NewBuildingArea", "Farmland", "Plough4", "Unuseland", "TargetBuildingArea4", "TargetPlough", "RealUseBuildingArea", "RealUsePlough" };

            //表格3 string[] FieldName = new string[] { "RegionName3", "RegionIndex3", "InvolvedTown", "PlanuseBuilding_all", "PlanusePlough_all", "RealuseBuilding_all", "RealusePlough_all", "PlanuseBuilding_setup", "PlanusePlough_setup", "RealuseBuilding_setup", "RealusePlough_setup", "PlanuseBuilding_out", "PlanusePlough_out", "RealuseBuilding_out", "RealusePlough_out","Planreturnbuildingarea","Planreturnplough","Realreturnbuildingarea","Realreturnplough" };
            //string[] FieldName = new string[] { "Index", "District", "RegionName", "RegionIndex", "AdmitedTime", "ImplementSpan", "ProcessState" };
            //string[] FieldName = new string[] { "Index", "District", "RegionName", "RegionIndex", "AdmitedTime", "ImplementSpan", "ProcessState" };

            for (int i = 0; i < FieldName.Length; i++)
            {
                pDataTable1.Columns.Add(FieldName[i]);
            }
            DataTable dtResult = pDT.Clone();
            DataTable dtName = pDT.DefaultView.ToTable(true, "XMBH");
            for (int i = 0; i < dtName.Rows.Count; i++)
            {
                DataRow[] rows = pDT.Select("XMBH='" + dtName.Rows[i][0] + "'");

                //temp用来存储筛选出来的数据
                DataTable temp = dtResult.Clone();
                foreach (DataRow row in rows)
                {
                    temp.Rows.Add(row.ItemArray);
                }
               
                DataRow[] jianxin_rows = temp.Select("DKLX='建新地块'");


                DataRow pRow = pDataTable1.NewRow();
                string[] StrRow = new string[12];
                StrRow[0] = rows[0]["区名"].ToString();
                StrRow[1] = rows[0]["XMMC"].ToString();
                StrRow[2] = jianxin_rows.Count().ToString();
                StrRow[3] = calcu_sum(jianxin_rows).ToString();

                pRow.ItemArray = StrRow;
                pDataTable1.Rows.Add(pRow);
            }
            dgv_Table4.DataSource = pDataTable1;
        }
        private void btnPreview5_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            pDT.DefaultView.Sort = "区名 desc";
            DataTable pDataTable1 = new DataTable();//建立一个table
            string[] FieldName = new string[] { "Index5", "ProjectName", "ReclamationProjectsNumber", "ApprovedPlotsNumber", "Area5", "AcceptanceItemsNumber", "AcceptancePlotsNumber", "ImplementationArea", "ReturnBuildingArea", "NewPloughArea", "ReturnPloughArea"};

            for (int i = 0; i < FieldName.Length; i++)
            {
                pDataTable1.Columns.Add(FieldName[i]);
            }
            DataTable dtResult = pDT.Clone();
            DataTable dtName = pDT.DefaultView.ToTable(true, "XMBH");
            for (int i = 0; i < dtName.Rows.Count; i++)
            {
                DataRow[] rows = pDT.Select("XMBH='" + dtName.Rows[i][0] + "'");

                //temp用来存储筛选出来的数据
                //DataTable temp = dtResult.Clone();
                //foreach (DataRow row in rows)
                //{
                //    temp.Rows.Add(row.ItemArray);
                //}

                //DataRow[] jianxin_rows = temp.Select("DKLX='建新地块'");


                DataRow pRow = pDataTable1.NewRow();
                string[] StrRow = new string[11];
                StrRow[0] = (i + 1).ToString();
                StrRow[1] = rows[0]["XMMC"].ToString();


                pRow.ItemArray = StrRow;
                pDataTable1.Rows.Add(pRow);
            }
            dgv_Table5.DataSource = pDataTable1;
        }
        private void btnExport1_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            Janus.Windows.Ribbon.ButtonCommand btn = (Janus.Windows.Ribbon.ButtonCommand)sender;
            export(btn);
        }

        private void btnExport2_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            Janus.Windows.Ribbon.ButtonCommand btn = (Janus.Windows.Ribbon.ButtonCommand)sender;
            export(btn);
        }

        private void btnExport3_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            Janus.Windows.Ribbon.ButtonCommand btn = (Janus.Windows.Ribbon.ButtonCommand)sender;
            export(btn);
        }

        private void btnExport4_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            Janus.Windows.Ribbon.ButtonCommand btn = (Janus.Windows.Ribbon.ButtonCommand)sender;
            export(btn);
        }

        private void btnExport5_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            Janus.Windows.Ribbon.ButtonCommand btn = (Janus.Windows.Ribbon.ButtonCommand)sender;
            export(btn);
        }

 


    }
}