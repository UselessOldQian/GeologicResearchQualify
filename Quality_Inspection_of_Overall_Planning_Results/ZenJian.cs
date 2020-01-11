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
            pErrorDataTable.Columns.Add("ErrorCode");
            pErrorDataTable.Columns.Add("ErrorLayer");
            pErrorDataTable.Columns.Add("ErrorField");
            pErrorDataTable.Columns.Add("ErrorObjectID");
            pErrorDataTable.Columns.Add("ErrorText");
            pErrorDataTable.Columns.Add("ErrorExcept");
            pErrorDataTable.Columns.Add("ErrorCheck");
            StatisticTable.Columns.Add("Indicator");
            StatisticTable.Columns.Add("IndicatorArea");
            StatisticTable.Columns.Add("MeasurementArea");
            StatisticTable.Columns.Add("DiffArea");
            treeView2.ExpandAll();

            //设置表格背景色
            dgvError.RowsDefaultCellStyle.BackColor = Color.Ivory;
            dgvTable.RowsDefaultCellStyle.BackColor = Color.Ivory;
            dgvStastic.RowsDefaultCellStyle.BackColor = Color.Ivory;

            //设置交替行的背景色
            dgvError.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;
            dgvTable.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;
            dgvStastic.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;
        }

        private void Initialize()
        {
            myDelegateAppendTextInfo = new AppendTextInfo(AppendTextInfoMethod);
            myDelegateUpdateBarValue = new UpdateBarValue(UpdateBarValueMethod);
            myDelegateUpdateUiStatus = new UpdateUiStatus(UpdateStatusBarMethod);
        }

        public void AppendTextInfoMethod(string strMsg)
        {
            if (null != InformationBox && !InformationBox.IsDisposed && strMsg != null)
            {
                InformationBox.AppendText(strMsg);
            }
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
                }
                pDT = LD.ShowTableInDataGridView_zenjian((ITable)axMapControl1.get_Layer(0), dgvTable, out FieldName);
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

        private void dgvTable_Scroll(object sender, ScrollEventArgs e)
        {

        }

        DataTable pErrorDataTable = new DataTable();
        public string[] NewLayersName = { "XZQ", "GHFW", "JQDLTB", "CSKFBJNGHYT", "JSYDKZX", "YJJBNT", "STKJKZX", "JSYDHJBNTGZ2035", "JLHDK" };
        public string[] ChineseLayerName = { "行政区", "规划范围", "基期地类图斑", "城市开发边界内规划用途", "建设用地控制线", "永久基本农田", "生态空间控制线", "建设用地和基本农田管制", "现状建设用地减量化地块" };
       //基本检查
        private void btnBasicCheck_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            if (cbIsClear.Checked == true)
            {
                pErrorDataTable.Rows.Clear();
            }
            ProcessBarMaxValue = NewLayersName.Length;
            progressBar1.Maximum = ProcessBarMaxValue;
            string AppendText = "\r\n基本检查\r\n时间:" + DateTime.Now.ToString();
            this.Invoke(this.myDelegateAppendTextInfo, new object[] { AppendText });
            int plusnum = (int)(progressBar1.Maximum / NewLayersName.LongLength);
            for (int i = 0; i < NewLayersName.Length; i++)
            {
                this.Invoke(this.myDelegateUpdateUiStatus, new object[] { "已完成" + (progressBar1.Value * 100 / progressBar1.Maximum).ToString() + "%" });
                ILayer layerresult = GetLayerByName(NewLayersName[i]);
                if (layerresult == null)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR1101:" + GetChineseName(NewLayersName[i]) + "不存在" });
                    pErrorDataTable.Rows.Add(new object[] { "1101", NewLayersName[i], null, null, GetChineseName(NewLayersName[i]) + "不存在", false, true });
                    continue;
                }
                IFeatureClass pFeaCls = (layerresult as IFeatureLayer).FeatureClass;
                //再通过IGeoDataset接口获取FeatureClass坐标系统
                ISpatialReference pSpatialRef = (pFeaCls as IGeoDataset).SpatialReference;
                if (pSpatialRef.Name.ToUpper() == "UNKNOWN")
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR2201:" + GetChineseName(NewLayersName[i]) + "投影为" + pSpatialRef.Name });
                    pErrorDataTable.Rows.Add(new object[] { "2201", NewLayersName[i], null, null, GetChineseName(NewLayersName[i]) + "投影为" + pSpatialRef.Name, false, true });
                }
                this.Invoke(this.myDelegateUpdateBarValue, new object[] { i });
            }
            this.Invoke(myDelegateUpdateBarValue, new object[] { 0 });
            BindingSource bind = new BindingSource();//绑定错误窗口的数据源
            bind.DataSource = pErrorDataTable;
            dgvError.DataSource = bind;
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "基本检查完成" });
            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\n基本检查完成\r\n" });
        }

        public string switchName(string name)
        {
            switch (name)
            {
                case "XZQ":
                    return "行政区";
                case "GHFW":
                    return "规划范围";
                case "JQDLTB":
                    return "基期地类图斑";
                case "CSKFBJNGHYT":
                    return "城市开发边界内规划用途";
                case "JSYDKZX":
                    return "建设用地控制线";
                case "STKJKZX":
                    return "生态空间控制线";
                case "YJJBNT":
                    return "永久基本农田";
                case "JSYDHJBNTGZ2035":
                    return "建设用地和基本农田管制";
                case "JLHDK":
                    return "减量化地块";
                case "TDLYJGTZB":
                    return "土地利用结构调整表";
                case "GDZBPHB":
                    return "耕地占补平衡表";
                case "CSKFBJ":
                    return "城市开发边界";
                default:
                    return "";
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

        #region 属性检查按钮
        private void btnAttributeCheck_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            if (cbIsClear.Checked == true)
            {
                pErrorDataTable.Rows.Clear();
            }
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "正在属性检查..." });
            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\n数据属性检查\r\n时间:" + DateTime.Now.ToString() });
            progressBar1.Maximum = 8;
            ILayer layerresult = GetLayerByName("XZQ");
            if (layerresult != null)
            {
                CheckAttributeBSM(layerresult);
                CheckAttributeYSDM(layerresult);
                CheckAttributeXZQDM(layerresult);
                CheckAttributeXZQMC(layerresult);
                CheckAttributeMJ(layerresult, "MJ");
                CheckAttributeMSorSM(layerresult, "MS");
                this.Invoke(myDelegateUpdateUiStatus, new object[] { "图层XZQ属性检查完成" });
            }
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });

            layerresult = GetLayerByName("JQDLTB");
            if (layerresult != null)
            {
                CheckAttributeBSM(layerresult);
                CheckAttributeYSDM(layerresult);
                CheckAttributeDLBM_SX(layerresult);
                CheckAttributeMJ(layerresult, "TBMJ");
                this.Invoke(myDelegateUpdateUiStatus, new object[] { "图层JQDLTB属性检查完成" });
            }
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });

            layerresult = GetLayerByName("CSKFBJNGHYT");
            if (layerresult != null)
            {
                CheckAttributeBSM(layerresult);
                CheckAttributeYSDM(layerresult);
                CheckAttributeXZQDM(layerresult);
                CheckAttributeXZQMC(layerresult);
                CheckAttributeMJ(layerresult, "MJ");
                CheckAttributeMSorSM(layerresult, "SM");
                this.Invoke(myDelegateUpdateUiStatus, new object[] { "图层CSKFBJNGHYT属性检查完成" });
            }
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });

            layerresult = GetLayerByName("JSYDKZX");
            if (layerresult != null)
            {
                CheckAttributeBSM(layerresult);
                CheckAttributeYSDM(layerresult);
                CheckAttributeXZQDM(layerresult);
                CheckAttributeXZQMC(layerresult);
                this.Invoke(myDelegateUpdateUiStatus, new object[] { "图层JSYDKZX属性检查完成" });
            }
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });

            layerresult = GetLayerByName("YJJBNT");
            if (layerresult != null)
            {
                CheckAttributeBSM(layerresult);
                CheckAttributeYSDM(layerresult);
                CheckAttributeXZQDM(layerresult);
                CheckAttributeXZQMC(layerresult);
                CheckAttributeSFCSZB(layerresult);
                CheckAttributeMJ(layerresult, "MJ");
                CheckAttributeMSorSM(layerresult, "SM");
                this.Invoke(myDelegateUpdateUiStatus, new object[] { "图层YJJBNT属性检查完成" });
            }
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });

            layerresult = GetLayerByName("STKJKZX");
            if (layerresult != null)
            {
                CheckAttributeBSM(layerresult);
                CheckAttributeYSDM(layerresult);
                CheckAttributeXZQDM(layerresult);
                CheckAttributeXZQMC(layerresult);
                CheckTextAttribute(layerresult, "BHLX", BHLXrange, 10);
                CheckTextAttribute(layerresult, "GKDJ", GKDJrange, 10);
                CheckAttributeMJ(layerresult, "MJ");
                CheckAttributeMSorSM(layerresult, "SM");
                this.Invoke(myDelegateUpdateUiStatus, new object[] { "图层STKJKZX属性检查完成" });
            }
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });

            layerresult = GetLayerByName("JSYDHJBNTGZ2035");
            if (layerresult != null)
            {
                CheckAttributeBSM(layerresult);
                CheckAttributeYSDM(layerresult);
                CheckAttributeXZQDM(layerresult);
                CheckAttributeXZQMC(layerresult);
                CheckTextAttribute(layerresult, "GZQLXDM", GZQLXDMrange, 3);
                CheckTextAttribute(layerresult, "GZQLXMC", GZQLXMCrange, 20);
                CheckAttributeMJ(layerresult, "GZQMJ");
                this.Invoke(myDelegateUpdateUiStatus, new object[] { "图层JSYDHJBNTGZ2035属性检查完成" });
            }
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });

            layerresult = GetLayerByName("JLHDK");
            if (layerresult != null)
            {
                CheckAttributeBSM(layerresult);
                CheckAttributeYSDM(layerresult);
                CheckTextAttribute(layerresult, "JQDLDM", JQDLDMrange, 3);
                CheckTextAttribute(layerresult, "JQDLMC", JQDLMCrange, 10);
                CheckAttributeMJ(layerresult, "QYMJ");
                CheckTextAttribute(layerresult, "SSSX", SSSXrange, 10);
                CheckAttributeMSorSM(layerresult, "SM");
                this.Invoke(myDelegateUpdateUiStatus, new object[] { "图层JLHDK属性检查完成" });
            }
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateUpdateBarValue, new object[] { 0 });
            BindingSource bind = new BindingSource();
            bind.DataSource = pErrorDataTable;
            dgvError.DataSource = bind;
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "属性检查完成" });
            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\n属性检查完成\r\n" });
        }

        private void CheckGHFWAttribute()
        {
            ILayer player = GetLayerByName("GHFW");
            if (player == null) { return; }
            ITable ptable = player as ITable;
            string[] attributes = { "GDBYL", "YJJBNTBHRW", "XZJSYDZGDMJ", "TDZZBCGD", "XZJSYDJLHMJ", "STBHHXMJ", "STKJMJ", "CSKFBJMJ", "JSYDZGM", "CSKFBJNXZJSYDMJ" };
            for (int i = 0; i < attributes.Length; i++)
            {
                int attr = ptable.FindField(attributes[i]);
                if (attr < 0)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName("GHFW") + "的属性字段" + attributes[i] + "不存在或正名命名错误" });
                    pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, attributes[i], null, GetChineseName("GHFW") + "的属性字段" + attributes[i] + "不存在或正名命名错误", false, true });
                    continue;
                }
                IField pfield = ptable.Fields.get_Field(attr);
                if (esriFieldType.esriFieldTypeDouble != pfield.Type)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段" + attributes[i] + "类型不是Double" });
                    pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, attributes[i], null, GetChineseName(player.Name) + "的属性字段" + attributes[i] + "类型不是Double", false, true });
                    continue;
                }
                ICursor pCursor = ptable.Search(null, false);
                IRow pRrow = pCursor.NextRow();
                while (pRrow != null)
                {
                    if (Convert.IsDBNull(pRrow.get_Value(attr)))
                    {
                        this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3601:" + GetChineseName(player.Name) + "的属性字段" + attributes[i] + " objectID=" + pRrow.OID + "的值为空" });
                        pErrorDataTable.Rows.Add(new object[] { "3601", player.Name, attributes[i], pRrow.OID, GetChineseName(player.Name) + "的属性字段" + attributes[i] + " objectID=" + pRrow.OID + "的值为空", false, true });
                    }
                    pRrow = pCursor.NextRow();
                }
            }
        }


        /// <summary>
        /// 检查BSM字段
        /// </summary>
        /// <param name="player">要检查的图层</param>
        private void CheckAttributeBSM(ILayer player)
        {
            ITable ptable = (ITable)player;
            int FieldIndex = ptable.FindField("BSM");
            if (FieldIndex < 0)
            {
                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段BSM不存在或正名命名错误" });
                pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "BSM", null, GetChineseName(player.Name) + "的属性字段BSM不存在或正名命名错误", false, true });
                return;
            }
            int IDIndex = ptable.FindField("OBJECTID");
            IField pfield = ptable.Fields.get_Field(FieldIndex);
            if (pfield != null)
            {
                if (esriFieldType.esriFieldTypeInteger != pfield.Type && esriFieldType.esriFieldTypeSmallInteger != pfield.Type)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段BSM类型不是int" });
                    pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "BSM", null, GetChineseName(player.Name) + "的属性字段BSM类型不是int", false, true });
                    return;
                }
                ICursor pCursor = ptable.Search(null, false);
                IRow pRrow = pCursor.NextRow();
                while (pRrow != null)
                {
                    if (!Convert.IsDBNull(pRrow.get_Value(FieldIndex)))
                    {
                        if (int.Parse(pRrow.get_Value(FieldIndex).ToString()) < 0)
                        {
                            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3401:" + GetChineseName(player.Name) + "的属性字段BSM objectID=" + pRrow.get_Value(IDIndex) + "的值不在值域内" });
                            pErrorDataTable.Rows.Add(new object[] { "3401", player.Name, "BSM", pRrow.get_Value(IDIndex).ToString(), GetChineseName(player.Name) + "的属性字段BSM objectID=" + pRrow.get_Value(IDIndex) + "的值不在值域内", false, true });
                        }
                    }
                    else
                    {
                        //this.invoke(myDelegateAppendTextInfo,new object[] {"\r\nERROR3601:" + player.Name + "的属性字段BSM objectID=" + pRrow.get_Value(IDIndex) + "的值为空"}); 
                    }
                    pRrow = pCursor.NextRow();
                }
            }
            else
            {
                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段BSM不存在或正名命名错误" });
                pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "BSM", null, GetChineseName(player.Name) + "的属性字段BSM不存在或正名命名错误", false, true });
                return;
            }
        }

        string[] YSDMrange = { "1000600000", "1000600100", "1000600200", "1000609000", 
                             "2003000000","2003010000","2003010100","2003020000",
                             "2003020100","2003020110","2003020120","2003020130",
                             "2003020140","2003020150","2003020200","2003020221",
                             "2003020231","2003020500","2003020510","2003030000",
                             "2003030200","2003030210","2003030220","2003030230",
                             "2003030240","2003030300","2003030301","2003030302",
                             "2003030303","2003030304","2003030305","2003030306",
                             "2003030307","2003030308","2003030309","2003030310",
                             "2003030311","2003030312","2003030313","2003030314",
                             "2003030315","2003030316","2003030317","2003030318",
                             "2003030319","2003030320","2003030600","2003039900",
                             "2003039910","2003039920","2003039930","2003039940",
                             "2003040000","2003040100","2003040200","2003050000",
                             "2003050100","2003050200","2003050300"};

        /// <summary>
        /// 检查YSDM字段值
        /// </summary>
        /// <param name="player">要检查的图层</param>
        private void CheckAttributeYSDM(ILayer player)
        {
            int index = 1;
            ITable ptable = (ITable)player;
            int FieldIndex = ptable.FindField("YSDM");
            if (FieldIndex < 0)
            {
                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段YSDM不存在或正名命名错误" });
                pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "YSDM", null, GetChineseName(player.Name) + "的属性字段YSDM不存在或正名命名错误", false, true });
                return;
            }
            int IDIndex = ptable.FindField("OBJECTID");
            IField pfield = ptable.Fields.get_Field(FieldIndex);
            if (pfield != null)
            {
                if (esriFieldType.esriFieldTypeString != pfield.Type)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段YSDM类型不是Text" });
                    pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "YSDM", null, GetChineseName(player.Name) + "的属性字段YSDM类型不是Text", false, true });
                    return;
                }
                if (pfield.Length != 10)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段YSDM字段长度不为10" });
                    pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "YSDM", null, GetChineseName(player.Name) + "的属性字段YSDM字段长度不为10", false, true });
                }
                ICursor pCursor = ptable.Search(null, false);
                IRow pRrow = pCursor.NextRow();
                while (pRrow != null)
                {
                    if (Convert.IsDBNull(pRrow.get_Value(FieldIndex)))
                    {
                        //this.Invoke(myDelegateAppendTextInfo,new object[] {"\r\nERROR3601:" + player.Name + "的属性字段YSDM objectID=" + pRrow.get_Value(IDIndex) + "的值为空"}); 
                    }
                    else
                    {
                        string YSDMValue = "";

                        switch (player.Name)
                        {
                            case "XZQ":
                            case "行政区":
                                YSDMValue = "1000600100";
                                break;
                            case "GHFW":
                            case "规划范围":
                                YSDMValue = "1000600200";
                                break;
                            case "JQDLTB":
                            case "基期地类图斑":
                                YSDMValue = "2003010100";
                                break;
                            case "CSKFBJNGHYT":
                            case "城市开发边界内规划用途":
                                YSDMValue = "2003020241";
                                break;
                            case "JSYDKZX":
                            case "建设用地控制线":
                                YSDMValue = "2003020140";
                                break;
                            case "STKJKZX":
                            case "生态空间控制线":
                                YSDMValue = "2003020120";
                                break;
                            case "YJJBNT":
                            case "永久基本农田":
                                YSDMValue = "2003020110";
                                break;
                            case "JSYDHJBNTGZ2035":
                            case "建设用地和基本农田管制":
                                YSDMValue = "2003020221";
                                break;
                            case "JLHDK":
                            case "减量化地块":
                                YSDMValue = "2003020510";
                                break;
                            default:
                                break;
                        }

                        if ((string)pRrow.get_Value(FieldIndex) != YSDMValue && YSDMValue != "")
                        {
                            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3301:" + GetChineseName(player.Name) + "的属性字段YSDM objectID=" + pRrow.get_Value(IDIndex) + "的值不符合要求" });
                            pErrorDataTable.Rows.Add(new object[] { "3301", player.Name, "YSDM", pRrow.get_Value(IDIndex).ToString(), GetChineseName(player.Name) + "的属性字段YSDM objectID=" + pRrow.get_Value(IDIndex) + "的值不符合要求", false, true });
                        }
                    }
                    pRrow = pCursor.NextRow();
                    index++;
                }
            }
            else
            {
                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段YSDM不存在或正名命名错误" });
                pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "YSDM", null, GetChineseName(player.Name) + "的属性字段YSDM不存在或正名命名错误", false, true });
                return;
            }
        }

        /// <summary>
        /// 检查XZQDM字段
        /// </summary>
        /// <param name="player">要检查的图层</param>
        private void CheckAttributeXZQDM(ILayer player)
        {
            ITable ptable = (ITable)player;
            int FieldIndex = ptable.FindField("XZQDM");
            if (FieldIndex < 0)
            {
                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段XZQDM不存在或正名命名错误" });
                pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "XZQDM", null, GetChineseName(player.Name) + "的属性字段XZQDM不存在或正名命名错误", false, true });
                return;
            }
            int IDIndex = ptable.FindField("OBJECTID");
            IField pfield = ptable.Fields.get_Field(FieldIndex);
            if (pfield != null)
            {
                if (esriFieldType.esriFieldTypeString != pfield.Type)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段XZQDM类型不是Text" });
                    pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "XZQDM", null, GetChineseName(player.Name) + "的属性字段XZQDM类型不是Text", false, true });
                    return;
                }
                if (pfield.Length != 12)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段XZQDM字段长度不为12" });
                    pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "XZQDM", null, GetChineseName(player.Name) + "的属性字段XZQDM字段长度不为12", false, true });
                }
                ICursor pCursor = ptable.Search(null, false);
                IRow pRrow = pCursor.NextRow();
                while (pRrow != null)
                {
                    if (Convert.IsDBNull(pRrow.get_Value(FieldIndex)))
                    {
                        this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3601:" + GetChineseName(player.Name) + "的属性字段XZQDM objectID=" + pRrow.get_Value(IDIndex) + "的值为空" });
                        pErrorDataTable.Rows.Add(new object[] { "3601", player.Name, "XZQDM", pRrow.get_Value(IDIndex).ToString(), GetChineseName(player.Name) + "的属性字段XZQDM objectID=" + pRrow.get_Value(IDIndex) + "的值为空", false, true });
                    }
                    pRrow = pCursor.NextRow();
                }
            }
            else
            {
                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段XZQDM不存在或正名命名错误" });
                pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "XZQDM", null, GetChineseName(player.Name) + "的属性字段XZQDM不存在或正名命名错误", false, true });
                return;
            }
        }

        /// <summary>
        /// 检查XZQMC字段
        /// </summary>
        /// <param name="player">要检查的图层</param>
        private void CheckAttributeXZQMC(ILayer player)
        {
            ITable ptable = (ITable)player;
            int FieldIndex = ptable.FindField("XZQMC");
            if (FieldIndex < 0)
            {
                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段XZQMC不存在或正名命名错误" });
                pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "XZQMC", null, GetChineseName(player.Name) + "的属性字段XZQMC不存在或正名命名错误", false, true });
                return;
            }
            int IDIndex = ptable.FindField("OBJECTID");
            IField pfield = ptable.Fields.get_Field(FieldIndex);
            if (pfield != null)
            {
                if (esriFieldType.esriFieldTypeString != pfield.Type)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段XZQMC类型不是Text" });
                    pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "XZQMC", null, GetChineseName(player.Name) + "的属性字段XZQMC类型不是Text", false, true });
                    return;
                }
                if (pfield.Length != 100)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段XZQMC字段长度不为100" });
                    pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "XZQMC", null, GetChineseName(player.Name) + "的属性字段XZQMC字段长度不为100", false, true });
                }
                ICursor pCursor = ptable.Search(null, false);
                IRow pRrow = pCursor.NextRow();
                while (pRrow != null)
                {
                    if (Convert.IsDBNull(pRrow.get_Value(FieldIndex)))
                    {
                        this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段XZQMC objectID=" + pRrow.get_Value(IDIndex) + "的值为空" });
                        pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "XZQMC", pRrow.get_Value(IDIndex).ToString(), GetChineseName(player.Name) + "的属性字段XZQMC objectID=" + pRrow.get_Value(IDIndex) + "的值为空", false, true });
                        //InformationBox.Text += "\r\nERROR3601:" + player.Name + "的属性字段XZQMC objectID=" + pRrow.get_Value(IDIndex) + "的值为空"; 
                    }
                    pRrow = pCursor.NextRow();
                }
            }
            else
            {
                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段XZQMC不存在或正名命名错误" });
                pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "XZQMC", null, GetChineseName(player.Name) + "的属性字段XZQMC不存在或正名命名错误", false, true });
                return;
            }
        }

        /// <summary>
        /// 检查面积字段
        /// </summary>
        /// <param name="player">要检查的图层</param>
        /// <param name="MJname">面积字段的名称</param>
        private void CheckAttributeMJ(ILayer player, string MJname)
        {
            ITable ptable = (ITable)player;
            int FieldIndex = ptable.FindField(MJname);
            if (FieldIndex < 0)
            {
                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段" + MJname + "不存在或正名命名错误" });
                pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, MJname, null, GetChineseName(player.Name) + "的属性字段" + MJname + "不存在或正名命名错误", false, true });
                return;
            }
            int IDIndex = ptable.FindField("OBJECTID");
            IField pfield = ptable.Fields.get_Field(FieldIndex);
            if (pfield != null)
            {
                if (esriFieldType.esriFieldTypeDouble != pfield.Type)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段" + MJname + "类型不是Double" });
                    pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, MJname, null, GetChineseName(player.Name) + "的属性字段" + MJname + "类型不是Double", false, true });
                    return;
                }
                ICursor pCursor = ptable.Search(null, false);
                IRow pRrow = pCursor.NextRow();
                while (pRrow != null)
                {
                    if (Convert.IsDBNull(pRrow.get_Value(FieldIndex)))
                    {
                        //this.Invoke(myDelegateAppendTextInfo,new object[] {"\r\nERROR3601:" + player.Name + "的属性字段" + MJname + " objectID=" + pRrow.get_Value(IDIndex) + "的值为空"}); 
                    }
                    else
                    {
                        {
                            if ((double)pRrow.get_Value(FieldIndex) < 0)
                            {
                                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3401:" + GetChineseName(player.Name) + "的属性字段" + MJname + " objectID=" + pRrow.get_Value(IDIndex) + "的值不在值域内" });
                                pErrorDataTable.Rows.Add(new object[] { "3401", player.Name, MJname, pRrow.get_Value(IDIndex).ToString(), GetChineseName(player.Name) + "的属性字段" + MJname + " objectID=" + pRrow.get_Value(IDIndex) + "的值不在值域内", false, true });
                            }
                        }
                    }
                    pRrow = pCursor.NextRow();
                }
            }
            else
            {
                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段" + MJname + "不存在或正名命名错误" });
                pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, MJname, null, GetChineseName(player.Name) + "的属性字段" + MJname + "不存在或正名命名错误", false, true });
                return;
            }
        }

        /// <summary>
        /// 检查说明或描述字段
        /// </summary>
        /// <param name="player">要检查的图层</param>
        /// <param name="MSorSM">字段名为描述还是说明</param>
        private void CheckAttributeMSorSM(ILayer player, string MSorSM)
        {
            int TextLength = 0;
            if (MSorSM == "MS")
            {
                TextLength = 100;
            }
            if (MSorSM == "SM")
            {
                TextLength = 200;
            }
            ITable ptable = (ITable)player;
            int FieldIndex = ptable.FindField(MSorSM);
            if (FieldIndex < 0)
            {
                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段" + MSorSM + "不存在或正名命名错误" });
                pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, MSorSM, null, GetChineseName(player.Name) + "的属性字段" + MSorSM + "不存在或正名命名错误", false, true });
                return;
            }
            int IDIndex = ptable.FindField("OBJECTID");
            IField pfield = ptable.Fields.get_Field(FieldIndex);
            if (pfield != null)
            {
                if (esriFieldType.esriFieldTypeString != pfield.Type)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段" + MSorSM + "类型不是Text" });
                    pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, MSorSM, null, GetChineseName(player.Name) + "的属性字段" + MSorSM + "类型不是Text", false, true });
                    return;
                }
                if (pfield.Length != TextLength)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段" + MSorSM + "字段长度不为100" });
                    pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, MSorSM, null, GetChineseName(player.Name) + "的属性字段" + MSorSM + "字段长度不为100", false, true });
                }
            }
            else
            {
                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段" + MSorSM + "不存在或正名命名错误" });
                pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, MSorSM, null, GetChineseName(player.Name) + "的属性字段" + MSorSM + "不存在或正名命名错误", false, true });
                return;
            }
        }

        /// <summary>
        /// 检查DLBM_SX字段
        /// </summary>
        /// <param name="player">要检查图层的名称</param>
        private void CheckAttributeDLBM_SX(ILayer player)
        {
            ITable ptable = (ITable)player;
            int FieldIndex = ptable.FindField("DLBM_SX");
            if (FieldIndex < 0)
            {
                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段DLBM_SX不存在或正名命名错误" });
                pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "DLBM_SX", null, GetChineseName(player.Name) + "的属性字段DLBM_SX不存在或正名命名错误", false, true });
                return;
            }
            int IDIndex = ptable.FindField("OBJECTID");
            IField pfield = ptable.Fields.get_Field(FieldIndex);
            if (pfield != null)
            {
                if (esriFieldType.esriFieldTypeString != pfield.Type)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段DLBM_SX类型不是Text" });
                    pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "DLBM_SX", null, GetChineseName(player.Name) + "的属性字段DLBM_SX类型不是Text", false, true });
                    return;
                }
                if (pfield.Length != 10)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段DLBM_SX字段长度不为10" });
                    pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "DLBM_SX", null, GetChineseName(player.Name) + "的属性字段DLBM_SX字段长度不为10", false, true });
                }
                ICursor pCursor = ptable.Search(null, false);
                IRow pRrow = pCursor.NextRow();
                while (pRrow != null)
                {
                    if (Convert.IsDBNull(pRrow.get_Value(FieldIndex)))
                    {
                        //this.Invoke(myDelegateAppendTextInfo,new object[] {"\r\nERROR3601:" + player.Name + "的属性字段DLBM_SX objectID=" + pRrow.get_Value(IDIndex) + "的值为空"}); 
                    }
                    pRrow = pCursor.NextRow();
                }
            }
            else
            {
                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段DLBM_SX不存在或正名命名错误" });
                pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "DLBM_SX", null, GetChineseName(player.Name) + "的属性字段DLBM_SX不存在或正名命名错误", false, true });
                return;
            }
        }

        string[] GHYTrange = { "城镇建设用地区", "产业基地", "产业社区", "战略预留区", "规划水域", "010", "021", "022", "030", "040" };
        string[] LXrange = { "城市开发边界内建设用地", "其他建设用地区" };


        string[] SFCSZBrange = { "Y", "N" };
        /// <summary>
        /// 检查SFCSZB字段
        /// </summary>
        /// <param name="player">要检查图层的名称</param>
        private void CheckAttributeSFCSZB(ILayer player)
        {
            ITable ptable = (ITable)player;
            int FieldIndex = ptable.FindField("SFCSZB");
            if (FieldIndex < 0)
            {
                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段SFCSZB不存在或正名命名错误" });
                pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "SFCSZB", null, GetChineseName(player.Name) + "的属性字段SFCSZB不存在或正名命名错误", false, true });
                return;
            }
            int IDIndex = ptable.FindField("OBJECTID");
            IField pfield = ptable.Fields.get_Field(FieldIndex);
            if (pfield != null)
            {
                if (esriFieldType.esriFieldTypeString != pfield.Type)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段SFCSZB类型不是Text" });
                    pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "SFCSZB", null, GetChineseName(player.Name) + "的属性字段SFCSZB类型不是Text", false, true });
                    return;
                }
                if (pfield.Length != 10)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段SFCSZB字段长度不为10" });
                    pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "SFCSZB", null, GetChineseName(player.Name) + "的属性字段SFCSZB字段长度不为10", false, true });
                }
                ICursor pCursor = ptable.Search(null, false);
                IRow pRrow = pCursor.NextRow();
                while (pRrow != null)
                {
                    if (Convert.IsDBNull(pRrow.get_Value(FieldIndex)))
                    {
                        //this.Invoke(myDelegateAppendTextInfo,new object[] {"\r\nERROR3601:" + player.Name + "的属性字段SFCSZB objectID=" + pRrow.get_Value(IDIndex) + "的值为空"}); 
                    }
                    else
                    {
                        if (SFCSZBrange.Contains((string)pRrow.get_Value(FieldIndex)) == false)
                        {
                            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3301:" + GetChineseName(player.Name) + "的属性字段SFCSZB objectID=" + pRrow.get_Value(IDIndex) + "的值不符合要求" });
                            pErrorDataTable.Rows.Add(new object[] { "3301", player.Name, "SFCSZB", pRrow.get_Value(IDIndex).ToString(), GetChineseName(player.Name) + "的属性字段SFCSZB objectID=" + pRrow.get_Value(IDIndex) + "的值不符合要求", false, true });
                        }
                    }
                    pRrow = pCursor.NextRow();
                }
            }
            else
            {
                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段SFCSZB不存在或正名命名错误" });
                pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "SFCSZB", null, GetChineseName(player.Name) + "的属性字段SFCSZB不存在或正名命名错误", false, true });
                return;
            }
        }

        string[] BHLXrange = { "110", "111", "112", "113", "114", "210", "211", "212", "213", "214", "215", "216", "217", "218", "219", "220", "221" };
        string[] GKDJrange = { "01", "02", "03", "04" };
        string[] GZQLXDMrange = { "01", "011", "012", "031", "033", "032", "040" };
        string[] GZQLXMCrange = { "允许建设区", "允许建设区(现状)", "允许建设区（现状）", "允许建设区(新增)", "允许建设区（新增）", "基本农田", "河湖水面", "一般农用地", "限制建设区", "禁止建设区" };
        string[] JQDLDMrange = { "20", "22", "25", "26", "27" };
        string[] JQDLMCrange = { "城镇建设用地", "工业仓储用地", "农村居民点用地", "交通运输用地", "其他建设用地" };
        string[] SSSXrange = { "近期", "远期" };
        /// <summary>
        /// 检查text类型的字段
        /// </summary>
        /// <param name="player">要检查图层的名称</param>
        /// <param name="TextAttributeName">要检查字段的名称</param>
        /// <param name="TextRange">允许的text的长度</param>
        /// <param name="TextLength">允许的text的值域</param>
        private void CheckTextAttribute(ILayer player, string TextAttributeName, string[] TextRange, int TextLength)
        {
            ITable ptable = (ITable)player;
            int FieldIndex = ptable.FindField(TextAttributeName);
            if (FieldIndex < 0)
            {
                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段" + TextAttributeName + "不存在或正名命名错误" });
                pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, TextAttributeName, null, GetChineseName(player.Name) + "的属性字段" + TextAttributeName + "不存在或正名命名错误", false, true });
                return;
            }
            int IDIndex = ptable.FindField("OBJECTID");
            IField pfield = ptable.Fields.get_Field(FieldIndex);
            if (pfield != null)
            {
                if (esriFieldType.esriFieldTypeString != pfield.Type)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段" + TextAttributeName + "类型不是Text" });
                    pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, TextAttributeName, null, GetChineseName(player.Name) + "的属性字段" + TextAttributeName + "类型不是Text", false, true });
                    return;
                }
                if (pfield.Length != TextLength)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段" + TextAttributeName + "字段长度不为" + TextLength.ToString() });
                    pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, TextAttributeName, null, GetChineseName(player.Name) + "的属性字段" + TextAttributeName + "字段长度不为" + TextLength.ToString(), false, true });
                }
                ICursor pCursor = ptable.Search(null, false);
                IRow pRrow = pCursor.NextRow();
                while (pRrow != null)
                {
                    if (Convert.IsDBNull(pRrow.get_Value(FieldIndex)))
                    {
                        this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3601:" + GetChineseName(player.Name) + "的属性字段" + TextAttributeName + " objectID=" + pRrow.get_Value(IDIndex) + "的值为空" });
                        pErrorDataTable.Rows.Add(new object[] { "3601", player.Name, TextAttributeName, pRrow.get_Value(IDIndex).ToString(), GetChineseName(player.Name) + "的属性字段" + TextAttributeName + " objectID=" + pRrow.get_Value(IDIndex) + "的值为空", false, true });
                    }
                    else
                    {
                        if (TextRange.Contains((string)pRrow.get_Value(FieldIndex)) == false)
                        {
                            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3301:" + GetChineseName(player.Name) + "的属性字段" + TextAttributeName + " objectID=" + pRrow.get_Value(IDIndex) + "的值不符合要求" });
                            pErrorDataTable.Rows.Add(new object[] { "3301", player.Name, TextAttributeName, pRrow.get_Value(IDIndex).ToString(), GetChineseName(player.Name) + "的属性字段" + TextAttributeName + " objectID=" + pRrow.get_Value(IDIndex) + "的值不符合要求", false, true });
                        }
                    }
                    pRrow = pCursor.NextRow();
                }
            }
            else
            {
                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段" + TextAttributeName + "不存在或正名命名错误" });
                pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, TextAttributeName, null, GetChineseName(player.Name) + "的属性字段" + TextAttributeName + "不存在或正名命名错误", false, true });
                return;
            }
        }
        #endregion

        CheckDataConsistent CDC = new CheckDataConsistent();
        private void btnDataConsistent_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            if (cbIsClear.Checked == true)
            {
                pErrorDataTable.Rows.Clear();
            }
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "正在上位规划落实情况检查..." });
            string AppendText = "\r\n上位规划落实情况检查\r\n时间:" + DateTime.Now.ToString();
            this.Invoke(this.myDelegateAppendTextInfo, new object[] { AppendText });
            progressBar1.Maximum = 5;
            ILayer layer1 = GetLayerByName("XZQ");
            ILayer layer2 = GetLayerByName("JQDLTB");
            ILayer layer3 = GetLayerByName("JSYDHJBNTGZ2035");
            ILayer[] Layerlist = { layer1, layer2, layer3 };
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeConsistent1(Layerlist, ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeConsistent2(GetLayerByName("行政区界") as IFeatureLayer, layer1 as IFeatureLayer, ref pErrorDataTable, "行政区界", "6401") });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("JBNTBHTB") as IFeatureLayer, GetLayerByName("YJJBNT") as IFeatureLayer, ref pErrorDataTable, "行政区界", "6401") });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("生态保护红线") as IFeatureLayer, GetLayerByName("STKJKZX") as IFeatureLayer, ref pErrorDataTable, "行政区界", "BHLX LIKE '11*'", "6401") });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CheckSpatial6401_5(GetLayerByName("CSKFBJNGHYT") as IFeatureLayer, GetLayerByName("城市开发边界") as IFeatureLayer, ref pErrorDataTable) });
            BindingSource bind = new BindingSource();
            bind.DataSource = pErrorDataTable;
            dgvError.DataSource = bind;
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "上位规划落实情况检查完毕" });
            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\n数据一致性检查完成\r\n完成时间:" + DateTime.Now.ToString() + "\r\n" });
            this.Invoke(myDelegateUpdateBarValue, new object[] { 0 });
        }

        public string CheckSpatial6401_5(IFeatureLayer Data1, IFeatureLayer Data2, ref DataTable pErrorDataTable)
        {
            if (Data2 == null) { return ""; }
            if (Data1 == null) { return ""; }
            ISpatialReference GRout = (Data1.FeatureClass as IGeoDataset).SpatialReference;
            ISpatialReference GRin = (Data2.FeatureClass as IGeoDataset).SpatialReference;
            if (GRout.SpatialReferenceImpl != GRin.SpatialReferenceImpl || GRout.Name != GRin.Name)
            { MessageBox.Show("外部数据" + Data1.Name + "与" + Data2.Name + "的坐标系不同，" + Data2.Name + "与" + Data1.Name + "范围无法比较，请修改"); return "\r\n外部数据" + Data1.Name + "与" + Data2.Name + "的坐标系不同"; }

            IGeometry Geo1 = CDC.UnionAllSelect(Data1.FeatureClass, null);
            IGeometry Geo2 = CDC.UnionAllSelect(Data2.FeatureClass, null);
            IGeometry XZQ_geo = CDC.UnionAllSelect((GetLayerByName("XZQ") as IFeatureLayer).FeatureClass, null);

            ITopologicalOperator pGeoInTP1 = Geo1 as ITopologicalOperator;
            IGeometry pDiff1 = pGeoInTP1.Difference(Geo2);
            IArea pArea1 = pDiff1 as IArea;
            ITopologicalOperator pGeoInTP2 = Geo2 as ITopologicalOperator;
            IGeometry pDiff2 = pGeoInTP2.Difference(Geo1);
            ITopologicalOperator pGeoInTP3 = pDiff2 as ITopologicalOperator;
            IGeometry pDiff3 = pGeoInTP3.Intersect(XZQ_geo, esriGeometryDimension.esriGeometry2Dimension);
            IArea pArea3 = pDiff3 as IArea;
            IFeatureLayer pFeatureLayerPlus = new FeatureLayerClass();
            pFeatureLayerPlus.FeatureClass = CT.CreateMemoryFeatureClass(Data1.FeatureClass);
            IFeature pFeaturePlus = pFeatureLayerPlus.FeatureClass.CreateFeature();
            pFeaturePlus.Shape = pDiff1;
            pFeaturePlus.Store();
            pFeatureLayerPlus.Name = "增加区域";

            IFeatureLayer pFeatureLayerMinus = new FeatureLayerClass();
            pFeatureLayerMinus.FeatureClass = CT.CreateMemoryFeatureClass(Data1.FeatureClass);
            IFeature pFeatureMinus = pFeatureLayerMinus.FeatureClass.CreateFeature();
            pFeatureMinus.Shape = pDiff3;
            pFeatureMinus.Store();
            pFeatureLayerMinus.Name = "减少区域";
            axMapControl1.AddLayer(pFeatureLayerPlus as IFeatureLayer);
            axMapControl1.AddLayer(pFeatureLayerMinus as IFeatureLayer);
            axMapControl1.Refresh();
            //IRelationalOperator RO = Geo2 as IRelationalOperator;
            //bool isEqual = RO.Within(Geo1);//.Relation(Geo1, relationDescription);
            if (Geo1 != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Geo1);
            if (Geo2 != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Geo2);
            if (pArea1.Area != 0 || pArea3.Area != 0)
            {
                pErrorDataTable.Rows.Add(new object[] { "6401", Data1.Name, null, null, GetChineseName(Data1.Name) + "与" + GetChineseName(Data2.Name) + "位于本行政区内范围不一致，" + GetChineseName(Data1.Name) + "图层增加面积为" + pArea1.Area + "平方米，减少面积为" + pArea3.Area + "平方米，不一致范围已在图上显示", false, true });
                return "\r\nERROR" + "6401" + ":" + GetChineseName(Data1.Name) + "与" + GetChineseName(Data2.Name) + "位于本行政区内范围不一致，" + GetChineseName(Data1.Name) + "图层增加面积为" + pArea1.Area + "平方米，减少面积为" + pArea3.Area + "平方米，不一致范围已在图上显示";
            }
            return null;
        }

        private void Check4301(ILayer player, string MJname)
        {
            if (player == null) { return; }
            ITable ptable = (ITable)player;
            if (ptable == null) { return; }
            int FieldIndex = ptable.FindField(MJname);
            if (FieldIndex < 0)
            {
                return;
            }
            int IDIndex = ptable.FindField("OBJECTID");
            IField pfield = ptable.Fields.get_Field(FieldIndex);
            if (pfield != null)
            {
                if (esriFieldType.esriFieldTypeDouble != pfield.Type)
                {
                    return;
                }
                ICursor pCursor = ptable.Search(null, false);
                IRow pRrow = pCursor.NextRow();
                while (pRrow != null)
                {
                    if (Convert.IsDBNull(pRrow.get_Value(FieldIndex)))
                    {
                        //this.Invoke(myDelegateAppendTextInfo,new object[] {"\r\nERROR3601:" + player.Name + "的属性字段" + MJname + " objectID=" + pRrow.get_Value(IDIndex) + "的值为空"}); 
                    }
                    else
                    {
                        if ((double)pRrow.get_Value(FieldIndex) == 0)
                        {
                            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR4301:" + GetChineseName(player.Name) + "碎多边形" + MJname + " objectID=" + pRrow.get_Value(IDIndex) + "面积小于4mm2" });
                            pErrorDataTable.Rows.Add(new object[] { "4301", player.Name, MJname, pRrow.get_Value(IDIndex).ToString(), GetChineseName(player.Name) + "碎多边形" + MJname + " objectID=" + pRrow.get_Value(IDIndex) + "面积小于4mm2", false, true });
                        }
                    }
                    pRrow = pCursor.NextRow();
                }
            }
        }

        #region 拓扑检查按钮
        CheckTopology CT = new CheckTopology();
        private void btnTopologyCheck_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "正在拓扑检查..." });
            if (cbIsClear.Checked == true)
            {
                pErrorDataTable.Rows.Clear();
            }
            progressBar1.Maximum = 32;
            string AppendText = "\r\n拓扑检查\r\n时间:" + DateTime.Now.ToString();
            this.Invoke(this.myDelegateAppendTextInfo, new object[] { AppendText });
            int plusnum = (int)(progressBar1.Maximum / NewLayersName.LongLength);
            string XZQ = "XZQ";
            string JQDLTB = "JQDLTB";
            string CSKFBJ = "CSKFBJ";
            string JSYDKZX = "JSYDKZX";
            string YJJBNT = "YJJBNT";
            string STKJKZX = "STKJKZX";
            string JSYDHJBNTGZ2035 = "JSYDHJBNTGZ2035";
            string JLHDK = "JLHDK";
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfIntersection(GetLayerByName(XZQ), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfIntersection(GetLayerByName(JQDLTB), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfIntersection(GetLayerByName(CSKFBJ), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfIntersection(GetLayerByName(JSYDKZX), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfIntersection(GetLayerByName(YJJBNT), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfIntersection(GetLayerByName(STKJKZX), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfIntersection(GetLayerByName(JSYDHJBNTGZ2035), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfIntersection(GetLayerByName(JLHDK), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSimple(GetLayerByName(XZQ), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSimple(GetLayerByName(JQDLTB), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSimple(GetLayerByName(CSKFBJ), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSimple(GetLayerByName(JSYDKZX), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSimple(GetLayerByName(YJJBNT), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSimple(GetLayerByName(STKJKZX), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSimple(GetLayerByName(JSYDHJBNTGZ2035), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSimple(GetLayerByName(JLHDK), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfOverlap(GetLayerByName(XZQ), ref pErrorDataTable, 1) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfOverlap(GetLayerByName(JQDLTB), ref pErrorDataTable, 100) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfOverlap(GetLayerByName(CSKFBJ), ref pErrorDataTable, 100) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfOverlap(GetLayerByName(JSYDKZX), ref pErrorDataTable, 100) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfOverlap(GetLayerByName(YJJBNT), ref pErrorDataTable, 100) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfOverlap(GetLayerByName(STKJKZX), ref pErrorDataTable, 1000) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfOverlap(GetLayerByName(JSYDHJBNTGZ2035), ref pErrorDataTable, 1000) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfOverlap(GetLayerByName(JLHDK), ref pErrorDataTable, 100) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            Check4301(GetLayerByName("XZQ"), "MJ");
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            Check4301(GetLayerByName("JQDLTB"), "TBMJ");
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            Check4301(GetLayerByName("CSKFBJNGHYT"), "MJ");
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            Check4301(GetLayerByName("JSYDKZX"), "MJ");
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            Check4301(GetLayerByName("YJJBNT"), "QYMJ");
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            Check4301(GetLayerByName("STKJKZX"), "QYMJ");
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            Check4301(GetLayerByName("JSYDHJBNTGZ2035"), "GZQMJ");
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            Check4301(GetLayerByName("JLHDK"), "QYMJ");
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            BindingSource bind = new BindingSource();
            bind.DataSource = pErrorDataTable;
            dgvError.DataSource = bind;
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "拓扑检查完毕" });
            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\n拓扑检查完成\r\n完成时间:" + DateTime.Now.ToString() + "\r\n" });
            this.Invoke(myDelegateUpdateBarValue, new object[] { 0 });
        }
        #endregion






        #region 导出为excle按钮
        private void btnErrorExport_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            if (dgvError.IsCurrentCellInEditMode == true)
            {
                dgvError.CurrentCell = null;
            }
            string filePath = "";
            SaveFileDialog s = new SaveFileDialog();
            s.Title = "保存Excel文件";
            s.Filter = "Excel文件(*.xlsx)|*.xlsx";
            s.FilterIndex = 1;
            if (s.ShowDialog() == DialogResult.OK)
            {
                filePath = s.FileName;
                string AppendText = "\r\n导出错误信息\r\n时间:" + DateTime.Now.ToString();
                this.Invoke(this.myDelegateAppendTextInfo, new object[] { AppendText });
                if (dgvError.Rows.Count <= 0)
                {
                    this.Invoke(this.myDelegateAppendTextInfo, new object[] { "\r\n提示：无数据导出" }); return;
                }
                DataTable tmpErrorDataTable = new DataTable("ErrorDT");
                DataTable modelTable = new DataTable("ModelTable");
                for (int column = 0; column < dgvError.Columns.Count; column++)
                {
                    if (dgvError.Columns[column].Visible == true)
                    {
                        DataColumn tempColumn = new DataColumn(dgvError.Columns[column].HeaderText, typeof(string));
                        tmpErrorDataTable.Columns.Add(tempColumn);

                        DataColumn modelColumn = new DataColumn(dgvError.Columns[column].Name, typeof(string));
                        modelTable.Columns.Add(modelColumn);
                    }
                }
                for (int row = 0; row < dgvError.Rows.Count; row++)
                {
                    if (Convert.IsDBNull(dgvError.Rows[row].Cells["ErrorCheck"].Value)) { continue; }
                    if (Convert.ToBoolean(dgvError.Rows[row].Cells["ErrorCheck"].Value) != true) { continue; }
                    DataRow tempRow = tmpErrorDataTable.NewRow();
                    for (int i = 0; i < tmpErrorDataTable.Columns.Count; i++)
                    {
                        if (i == 5)
                        {
                            if (Convert.IsDBNull(dgvError.Rows[row].Cells[modelTable.Columns[i].ColumnName].Value))
                            {
                                tempRow[i] = "否";
                                continue;
                            }
                            if (Convert.ToBoolean(dgvError.Rows[row].Cells[modelTable.Columns[i].ColumnName].Value) != true)
                            {
                                tempRow[i] = "否";
                                continue;
                            }
                            else
                            {
                                tempRow[i] = "是";
                                continue;
                            }
                        }
                        tempRow[i] = dgvError.Rows[row].Cells[modelTable.Columns[i].ColumnName].Value;
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



        //private void EndReport()
        //{
        //    object missing = System.Reflection.Missing.Value;
        //    try
        //    {
        //        //m_xlApp.Workbooks.Close();  
        //        //m_xlApp.Workbooks.Application.Quit();  
        //        //m_xlApp.Application.Quit();  
        //        //m_xlApp.Quit();  
        //    }
        //    catch { }
        //    finally
        //    {
        //        try
        //        {
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(m_xlApp.Workbooks);
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(m_xlApp.Application);
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(m_xlApp);
        //            m_xlApp = null;
        //        }
        //        catch { }
        //        try
        //        {
        //            //清理垃圾进程  
        //            this.killProcessThread();
        //        }
        //        catch { }
        //        GC.Collect();
        //    }
        //}
        #endregion
        private void btnLoadDirectory_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            //确认根目录文件夹
            FolderBrowserDialog folder = new FolderBrowserDialog();
            folder.Description = "选择所有文件存放目录";
            if (folder.ShowDialog() == DialogResult.OK)
            {
                string folderpath = folder.SelectedPath;
                search(folderpath);
            }
        }

        private void search(string filterstr)
        {
            //创建DirectoryInfo对象
            DirectoryInfo dir = new DirectoryInfo(filterstr);
            FileSystemInfo[] fs = dir.GetFileSystemInfos();

            //获取目录中文件和子目录
            foreach (FileSystemInfo fi in fs)   //FileSystemInfo类为FileInfo和DirectoryInfo对象提供基类。
            {
                if (fi.Attributes == FileAttributes.Directory)
                { //判断是否目录
                    search(fi.FullName);
                }
                else
                {
                    if (fi.Extension == ".mdb")          //搜索条件
                    {
                        OpenMDB(fi.FullName);
                    }
                }
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

        private void buttonCommand5_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "正在规划布局检查..." });
            if (cbIsClear.Checked == true)
            {
                pErrorDataTable.Rows.Clear();
            }
            progressBar1.Maximum = 5;
            string AppendText = "\r\n规划布局检查\r\n时间:" + DateTime.Now.ToString();
            this.Invoke(this.myDelegateAppendTextInfo, new object[] { AppendText });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("JSYDHJBNTGZ2035") as IFeatureLayer,GetLayerByName("CSKFBJNGHYT") as IFeatureLayer,ref pErrorDataTable,
                "JSYDHJBNTGZ2035","GZQLXDM LIKE '01*'","CSKFBJNGHYT","GHYT LIKE '010' OR GHYT LIKE '021' OR GHYT LIKE '022' OR GHYT LIKE '030'","6501"  )});
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("JSYDHJBNTGZ2035") as IFeatureLayer,GetLayerByName("CSKFBJNGHYT") as IFeatureLayer,ref pErrorDataTable,
                "JSYDHJBNTGZ2035","GZQLXDM LIKE '033'", "CSKFBJNGHYT","GHYT LIKE '040'","6501"  )});
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("CSKFBJNGHYT") as IFeatureLayer, GetLayerByName("JSYDKZX") as IFeatureLayer, ref pErrorDataTable, null, "LX LIKE '城市开发边界内建设用地'", "6502") });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("JSYDKZX") as IFeatureLayer, GetLayerByName("CSKFBJNGHYT") as IFeatureLayer, ref pErrorDataTable, "JSYDKZX", "LX LIKE '其他建设用地区'", "CSKFBJNGHYT", null, "RELATE(G1, G2, 'FF*F*****')", "6502") });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("JSYDKZX") as IFeatureLayer, GetLayerByName("JSYDHJBNTGZ2035") as IFeatureLayer, ref pErrorDataTable, "JSYDKZX", null, "JSYDHJBNTGZ2035", "GZQLXDM LIKE '01*'", "RELATE(G1, G2, 'T*F*T*F**')", "6502") });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeNotWithin(GetLayerByName("YJJBNT") as IFeatureLayer, GetLayerByName("JSYDKZX") as IFeatureLayer, ref pErrorDataTable, null) });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeNotWithin(GetLayerByName("STKJKZX") as IFeatureLayer, GetLayerByName("JSYDKZX") as IFeatureLayer, ref pErrorDataTable, "BHLX LIKE '11*'") });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });

            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("STKJKZX") as IFeatureLayer, GetLayerByName("JSYDHJBNTGZ2035") as IFeatureLayer, ref pErrorDataTable, "STKJKZX", "BHLX LIKE '11*'", "JSYDHJBNTGZ2035", "GZQLXDM LIKE '040'", "RELATE(G1, G2, 'T*F*T*F**')", "6503") });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("STKJKZX") as IFeatureLayer, GetLayerByName("YJJBNT") as IFeatureLayer, ref pErrorDataTable, "STKJKZX", "BHLX LIKE '210'", "YJJBNT", null, "RELATE(G1, G2, 'T*F*T*F**')", "6503") });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("CSKFBJNGHYT") as IFeatureLayer, GetLayerByName("STKJKZX") as IFeatureLayer, ref pErrorDataTable, "CSKFBJNGHYT", null, "STKJKZX", "GKDJ LIKE '04'", "6503") });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeNotWithin6503(GetLayerByName("CSKFBJNGHYT") as IFeatureLayer, GetLayerByName("STKJKZX") as IFeatureLayer, ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("YJJBNT") as IFeatureLayer, GetLayerByName("JSYDHJBNTGZ2035") as IFeatureLayer, ref pErrorDataTable, "YJJBNT", null, "JSYDHJBNTGZ2035", "GZQLXDM LIKE '031'", "RELATE(G1, G2, 'T*F*T*F**')", "6504") });

            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("JSYDHJBNTGZ2035") as IFeatureLayer,GetLayerByName("JLHDK") as IFeatureLayer,ref pErrorDataTable,
                "JSYDHJBNTGZ2035","GZQLXDM LIKE '032' OR GZQLXDM LIKE '033'","JLHDK",null,"6505"  )});
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("JQDLTB") as IFeatureLayer,GetLayerByName("JLHDK") as IFeatureLayer,ref pErrorDataTable,
                "JQDLTB","DLBM_SX LIKE '2*'","JLHDK",null,"6505"  )});
            BindingSource bind = new BindingSource();
            bind.DataSource = pErrorDataTable;
            dgvError.DataSource = bind;
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "规划布局检查完毕" });
            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\n规划布局检查完成\r\n完成时间:" + DateTime.Now.ToString() + "\r\n" });
            this.Invoke(myDelegateUpdateBarValue, new object[] { 0 });
        }

        DataTable StatisticTable = new DataTable();
        private void buttonCommand6_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            this.StatisticTable.Rows.Clear();
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "正在建设用地简化量规模统计..." });
            progressBar1.Maximum = 10;
            string AppendText = "\r\n建设用地简化量规模统计\r\n时间:" + DateTime.Now.ToString();
            this.Invoke(this.myDelegateAppendTextInfo, new object[] { AppendText });
            double[] MeasureArea = new double[10] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            MeasureArea[1] = CDC.getLayerArea(GetLayerByName("YJJBNT"), null);
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            MeasureArea[2] = CDC.getIntersectArea(GetLayerByName("JSYDHJBNTGZ2035"), GetLayerByName("JQDLTB"), "GZQLXDM LIKE '012'", "DLBM_SX LIKE '11*'");
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            MeasureArea[4] = CDC.getLayerArea(GetLayerByName("JLHDK"), null);
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            MeasureArea[5] = CDC.getLayerArea(GetLayerByName("STKJKZX"), "GKDJ LIKE '01' OR GKDJ LIKE '02'");
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            MeasureArea[6] = CDC.getLayerArea(GetLayerByName("STKJKZX"), null);
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            MeasureArea[7] = CDC.getLayerArea(GetLayerByName("CSKFBJNGHYT"), null);
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            MeasureArea[8] = CDC.getLayerArea(GetLayerByName("JSYDHJBNTGZ2035"), "GZQLXDM LIKE '01*'");
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            MeasureArea[9] = CDC.getIntersectArea(GetLayerByName("CSKFBJNGHYT"), GetLayerByName("JSYDHJBNTGZ2035"), "GHYT LIKE '030' OR GHYT LIKE '010' OR GHYT LIKE '021' OR GHYT LIKE '022'", "GZQLXDM LIKE '012'");
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            double[] IndicatorArea = CDC.StatisticalScaleGHFW(GetLayerByName("GHFW"));
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            string[] Attrs ={"耕地保有量（公顷）","永久基本农田保护任务（公顷）","新增建设用地占用耕地面积（公顷）","土地整治补充耕地（公顷）",
                               "现状建设用地减量化面积（公顷）","生态保护红线（一二类生态空间）面积（公顷）","生态空间面积（公顷）","城市开发边界面积（公顷）",
                           "建设用地总规模（公顷）","城市开发边界内新增建设用地面积（公顷）"};
            for (int i = 0; i < Attrs.Length; i++)
            {
                if (i == 0 || i == 3) { StatisticTable.Rows.Add(new object[] { Attrs[i], IndicatorArea[i].ToString(), "/", "/" }); continue; }
                StatisticTable.Rows.Add(new object[] { Attrs[i], IndicatorArea[i].ToString(), MeasureArea[i].ToString("0.00"), (IndicatorArea[i] - MeasureArea[i]).ToString("0.00") });
            }
            BindingSource bind = new BindingSource();
            bind.DataSource = StatisticTable;
            dgvStastic.DataSource = bind;
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "建设用地简化量规模统计完毕" });
            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\n建设用地简化量规模统计完成\r\n完成时间:" + DateTime.Now.ToString() + "\r\n" });
            this.uiTab1.SelectedTab = this.uiTab1.TabPages[2];
            this.Invoke(myDelegateUpdateBarValue, new object[] { 0 });
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

        private void btnCheckAll_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            cbIsClear.Checked = false;
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "正在所有检查..." });
            string AppendText = "\r\n***************************************************************************************************************************************************************************\r\n所有检查\r\n时间:" + DateTime.Now.ToString();
            this.Invoke(this.myDelegateAppendTextInfo, new object[] { AppendText });
            pErrorDataTable.Rows.Clear();
            this.btnBasicCheck_Click(sender, e);
            this.btnAttributeCheck_Click(sender, e);
            this.btnDataConsistent_Click(sender, e);
            this.btnTopologyCheck_Click(sender, e);
            this.buttonCommand5_Click(sender, e);
            this.buttonCommand1_Click(sender, e);
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "所有检查完毕" });
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "所有检查完毕" });
            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\n所有检查完成\r\n完成时间:" + DateTime.Now.ToString() + "\r\n***************************************************************************************************************************************************************************\r\n" });
            this.Invoke(myDelegateUpdateBarValue, new object[] { 0 });
            cbIsClear.Checked = true;
        }

        private void dgvError_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {

            string ColumnName = this.dgvError.Columns[e.ColumnIndex].Name;
            if (ColumnName == "ErrorCheck" || ColumnName == "ErrorExcept")
            {
                if (dgvError.IsCurrentCellInEditMode == true)
                {
                    dgvError.CurrentCell = null;
                }
                if (Convert.IsDBNull(dgvError.Rows[0].Cells[e.ColumnIndex].Value))
                {
                    for (int i = 0; i < dgvError.Rows.Count; i++)
                    {
                        dgvError.Rows[i].Cells[e.ColumnIndex].Selected = false;
                        dgvError.Rows[i].Cells[e.ColumnIndex].Value = true;
                    }
                }
                else if (Convert.ToBoolean(dgvError.Rows[0].Cells[e.ColumnIndex].Value) == true)
                {
                    for (int i = 0; i < dgvError.Rows.Count; i++)
                    {
                        dgvError.Rows[i].Cells[e.ColumnIndex].Selected = false;
                        dgvError.Rows[i].Cells[e.ColumnIndex].Value = false;
                    }
                }
                else
                {
                    for (int i = 0; i < dgvError.Rows.Count; i++)
                    {
                        dgvError.Rows[i].Cells[e.ColumnIndex].Selected = false;
                        dgvError.Rows[i].Cells[e.ColumnIndex].Value = true;
                    }
                }

            }
        }

        private void treeView2_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            dgvError.Sort(dgvError.Columns[0], ListSortDirection.Ascending);
            for (int i = 0; i < dgvError.Rows.Count; i++)
            {
                dgvError.Rows[i].Selected = false;
            }
            bool Firstflag = true;
            for (int i = 0; i < dgvError.Rows.Count; i++)
            {
                if (dgvError.Rows[i].Cells[0].Value.ToString() == treeView2.SelectedNode.Name)
                {
                    dgvError.Rows[i].Selected = true;
                    if (Firstflag == true) { dgvError.FirstDisplayedScrollingRowIndex = i; Firstflag = false; }
                }
            }
        }

        private void treeView2_BeforeCollapse(object sender, TreeViewCancelEventArgs e)
        {
            e.Cancel = true;
        }

        private void btnAllSelectExcept_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            if (dgvError.SelectedRows.Count <= 0) { return; }

            if (dgvError.IsCurrentCellInEditMode == true)
            {
                dgvError.CurrentCell = null;
            }
            if (Convert.IsDBNull(dgvError.SelectedRows[0].Cells[5].Value))
            {
                for (int i = 0; i < dgvError.SelectedRows.Count; i++)
                {
                    dgvError.SelectedRows[i].Cells[5].Value = true;
                }
            }
            else if (Convert.ToBoolean(dgvError.SelectedRows[0].Cells[5].Value) == true)
            {
                for (int i = 0; i < dgvError.SelectedRows.Count; i++)
                {
                    dgvError.SelectedRows[i].Cells[5].Value = false;
                }
            }
            else
            {
                for (int i = 0; i < dgvError.SelectedRows.Count; i++)
                {
                    dgvError.SelectedRows[i].Cells[5].Value = true;
                }
            }
        }

        private void btnAllSelectExport_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            if (dgvError.SelectedRows.Count <= 0) { return; }

            if (dgvError.IsCurrentCellInEditMode == true)
            {
                dgvError.CurrentCell = null;
            }
            if (Convert.IsDBNull(dgvError.SelectedRows[0].Cells[6].Value))
            {
                for (int i = 0; i < dgvError.SelectedRows.Count; i++)
                {
                    dgvError.SelectedRows[i].Cells[6].Value = true;
                }
            }
            else if (Convert.ToBoolean(dgvError.SelectedRows[0].Cells[6].Value) == true)
            {
                for (int i = 0; i < dgvError.SelectedRows.Count; i++)
                {
                    dgvError.SelectedRows[i].Cells[6].Value = false;
                }
            }
            else
            {
                for (int i = 0; i < dgvError.SelectedRows.Count; i++)
                {
                    dgvError.SelectedRows[i].Cells[6].Value = true;
                }
            }
        }

        private void btnLoadFixMDB_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            try
            {
                System.Windows.Forms.OpenFileDialog openShipFileDlg = new System.Windows.Forms.OpenFileDialog();
                openShipFileDlg.Filter = "新市镇MDB文件(*.mdb)|*.mdb";
                openShipFileDlg.Multiselect = false;
                openShipFileDlg.Title = "选择新市镇MDB文件";
                DialogResult dr = openShipFileDlg.ShowDialog();
                if (dr == DialogResult.OK)
                {
                    string strFullPath = openShipFileDlg.FileName;
                    if (strFullPath == "") return;
                    OpenFixMDB(strFullPath);
                    this.tabMapTableView.SelectedTab = tabMapTableView.TabPages[0];
                    uiStatusBar1.Panels[0].Text = "数据库读取完成";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void OpenFixMDB(string strFullPath)
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
                if (pDataset is IFeatureClass)
                {
                    if (!NewLayersName.Contains(pDataset.Name))
                    {
                        pDataset = pEnumDataset.Next();
                        continue;
                    }
                    pFeatureWorkspace = (IFeatureWorkspace)pWorkspace;
                    pFeatureLayer = new FeatureLayerClass();
                    pFeatureLayer.FeatureClass = pFeatureWorkspace.OpenFeatureClass(pDataset.Name);
                    pFeatureLayer.Name = pFeatureLayer.FeatureClass.AliasName;
                    axMapControl1.Map.AddLayer(pFeatureLayer);
                    axMapControl1.ActiveView.FocusMap.get_Layer(0).Visible = false;
                    axMapControl1.ActiveView.Refresh();
                    this.uiTab2.SelectedTab = uiTab2.TabPages[0];
                }
                else if (pDataset is ITable)
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

        private void dgvError_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (Convert.IsDBNull(dgvError.CurrentRow.Cells[3].Value))
            {
                return;
            }
            else
            {
                int OID = int.Parse(dgvError.CurrentRow.Cells[3].Value.ToString());
                ILayer player = GetLayerByName(dgvError.CurrentRow.Cells[1].Value.ToString());
                IFeatureSelection pFeatureSelection = (player as IFeatureLayer) as IFeatureSelection;
                //创建过滤器
                IQueryFilter pQueryFilter = new QueryFilterClass();
                //设置过滤器对象的查询条件
                pQueryFilter.WhereClause = "OBJECTID = " + OID.ToString();
                //根据查询条件选择要素
                pFeatureSelection.SelectFeatures(pQueryFilter, esriSelectionResultEnum.esriSelectionResultNew, false);
                ISimpleFillSymbol SFS = new SimpleFillSymbolClass();
                ISimpleLineSymbol ILS = new SimpleLineSymbolClass();
                SFS.Style = esriSimpleFillStyle.esriSFSSolid;
                SFS.Color = getRGB(255, 0, 0);
                ILS.Color = getRGB(0, 255, 0);
                ILS.Style = esriSimpleLineStyle.esriSLSSolid;
                ILS.Width = 13;
                SFS.Outline = ILS;
                pFeatureSelection.SelectionSymbol = SFS as ISymbol;
                IArea pArea = (player as IFeatureLayer).FeatureClass.GetFeature(OID).Shape as IArea;
                IPoint iPnt = pArea.LabelPoint;
                axMapControl1.Extent = (player as IFeatureLayer).FeatureClass.GetFeature(OID).Shape.Envelope;
                axMapControl1.CenterAt(iPnt);
            
                axMapControl1.Refresh();
            }
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

       

        private void btnTextAuxiliaryCheck_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            if (cbIsClear.Checked == true)
            {
                pErrorDataTable.Rows.Clear();
            }
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "正在文本辅助检查..." });
            string AppendText = "\r\n文本辅助检查\r\n时间:" + DateTime.Now.ToString();
            this.Invoke(this.myDelegateAppendTextInfo, new object[] { AppendText });
            progressBar1.Maximum = 6;
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckArea(GetLayerByName("XZQ"), null, treeView1, "TDLYJGTZB", "SSJD = '基准年'", "DLMJ", 2, ref pErrorDataTable, "6403") });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckArea("TDLYJGTZB","SSJD LIKE '规划年' AND ( DLMC LIKE '林地' OR DLMC LIKE '设施农业用地' OR DLMC LIKE '河湖水面')","DLMJ",treeView1,
                "TDLYJGTZB","SSJD LIKE '基准年' AND ( DLMC LIKE '林地' OR DLMC LIKE '设施农业用地' OR DLMC LIKE '河湖水面')","DLMJ",1,ref pErrorDataTable,"6403")});
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckArea("TDLYJGTZB","SSJD LIKE '规划年' AND ( DLMC LIKE '耕地')","DLMJ",treeView1,
                GetLayerByName("GHFW"),null,"GDBYL",1,ref pErrorDataTable,"6403")});
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckArea("TDLYJGTZB","SSJD LIKE '基准年'","DLMJ",treeView1,
                "TDLYJGTZB","SSJD LIKE '基准年'","DLMJ",2,ref pErrorDataTable,"6403")});
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckArea("GDZBPHB", "BZ LIKE '规划期间补充耕地'", "ZBMJ", treeView1, "GDZBPHB", "BZ LIKE '占用'", "ZBMJ", 1, ref pErrorDataTable, "6403") });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.JudgeArea(CDC.getArea("TDLYJGTZB","SSJD LIKE '规划年' AND ( DLMC LIKE '耕地')","DLMJ",treeView1)-CDC.getArea("TDLYJGTZB","SSJD LIKE '基准年' AND ( DLMC LIKE '耕地')","DLMJ",treeView1),"TDLYJGTZB",
                CDC.getArea("GDZBPHB","BZ LIKE '规划期间补充耕地'","ZBMJ",treeView1)-CDC.getArea("GDZBPHB","BZ LIKE '规划期间减少耕地'","ZBMJ",treeView1),"GDZBPHB",2,ref pErrorDataTable,"6403")});
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            BindingSource bind = new BindingSource();
            bind.DataSource = pErrorDataTable;
            dgvError.DataSource = bind;
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "文本辅助检查完毕" });
            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\n文本辅助检查完成\r\n完成时间:" + DateTime.Now.ToString() + "\r\n" });
            this.Invoke(myDelegateUpdateBarValue, new object[] { 0 });
        }
        //导出统计数据
        private void btnStatisticExport_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            if (dgvStastic.IsCurrentCellInEditMode == true)
            {
                dgvStastic.CurrentCell = null;
            }
            string filePath = "";
            SaveFileDialog s = new SaveFileDialog();
            s.Title = "保存Excel文件";
            s.Filter = "Excel文件(*.xlsx)|*.xlsx";
            s.FilterIndex = 1;
            if (s.ShowDialog() == DialogResult.OK)
            {
                filePath = s.FileName;
                string AppendText = "\r\n导出统计表格\r\n时间:" + DateTime.Now.ToString();
                this.Invoke(this.myDelegateAppendTextInfo, new object[] { AppendText });
                if (dgvStastic.Rows.Count <= 0)
                {
                    this.Invoke(this.myDelegateAppendTextInfo, new object[] { "\r\n提示：无数据导出" }); return;
                }
                DataTable tmpStatisticDataTable = new DataTable("StatisticDT");
                DataTable modelTable = new DataTable("ModelTable");
                for (int column = 0; column < dgvStastic.Columns.Count; column++)
                {
                    if (dgvStastic.Columns[column].Visible == true)
                    {
                        DataColumn tempColumn = new DataColumn(dgvStastic.Columns[column].HeaderText, typeof(string));
                        tmpStatisticDataTable.Columns.Add(tempColumn);

                        DataColumn modelColumn = new DataColumn(dgvStastic.Columns[column].Name, typeof(string));
                        modelTable.Columns.Add(modelColumn);
                    }
                }
                for (int row = 0; row < dgvStastic.Rows.Count; row++)
                {
                    DataRow tempRow = tmpStatisticDataTable.NewRow();
                    for (int i = 0; i < tmpStatisticDataTable.Columns.Count; i++)
                    {
                        tempRow[i] = dgvStastic.Rows[row].Cells[modelTable.Columns[i].ColumnName].Value;
                    }
                    tmpStatisticDataTable.Rows.Add(tempRow);
                }
                if (tmpStatisticDataTable == null)
                {
                    return;
                }
                //第二步：导出dataTable到Excel  
                long rowNum = tmpStatisticDataTable.Rows.Count;//行数  
                int columnNum = tmpStatisticDataTable.Columns.Count;//列数  
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
                        datas[0, i] = tmpStatisticDataTable.Columns[i].Caption;
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
                            object obj = tmpStatisticDataTable.Rows[r][tmpStatisticDataTable.Columns[i].ToString()];
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

        private void buttonCommand1_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            if (cbIsClear.Checked == true)
            {
                pErrorDataTable.Rows.Clear();
            }
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "正在图数一致性检查..." });
            string AppendText = "\r\n图数一致性检查\r\n时间:" + DateTime.Now.ToString();
            this.Invoke(this.myDelegateAppendTextInfo, new object[] { AppendText });
            progressBar1.Maximum = 4;
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckArea(GetLayerByName("JSYDHJBNTGZ2035"), "GZQLXDM LIKE '01*'", GetLayerByName("GHFW"), null, "JSYDZGM", 0, ref pErrorDataTable, "6402") });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckArea(GetLayerByName("YJJBNT"), null, GetLayerByName("GHFW"), null, "YJJBNTBHRW", 1, ref pErrorDataTable, "6402") });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckArea(GetLayerByName("CSKFBJNGHYT"), null, GetLayerByName("GHFW"), null, "CSKFBJMJ", 2, ref pErrorDataTable, "6402") });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckArea(GetLayerByName("STKJKZX"), "GKDJ LIKE '01' OR GKDJ LIKE '02'", GetLayerByName("GHFW"), null, "STBHHXMJ", 2, ref pErrorDataTable, "6402") });
            BindingSource bind = new BindingSource();
            bind.DataSource = pErrorDataTable;
            dgvError.DataSource = bind;
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "图数一致性检查完毕" });
            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\n图数一致性检查完成\r\n完成时间:" + DateTime.Now.ToString() + "\r\n" });
            this.Invoke(myDelegateUpdateBarValue, new object[] { 0 });
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

        private void dgvError_DataSourceChanged(object sender, EventArgs e)
        {
            string[] ErrorNumber = new string[] { "1101", "2201", "3201", "3301", "3401", "3601", "4301", "4101", "6401", "6402", "6403", "6501", "6502", "6503", "6504", "6505" };
            int[] ErrorIndex = new int[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            string[] ErrorMassage = new string[] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
            if (dgvError.Rows.Count == 0)
            {
                for (int i = 0; i < ErrorNumber.Length; i++)
                {
                    ErrorMassage[i] = ErrorNumber[i] + "(" + ErrorIndex[i].ToString() + "条)";
                }
                for (int i = 0; i < treeView2.Nodes[0].Nodes.Count; i++)
                {
                    treeView2.Nodes[0].Nodes[i].Text = ErrorMassage[i];
                }
                return;
            }
            for (int rows = 0; rows < dgvError.RowCount; rows++)
            {
                for (int i = 0; i < ErrorNumber.Length; i++)
                {
                    if (ErrorNumber[i] == dgvError.Rows[rows].Cells[0].Value.ToString())
                    {
                        ErrorIndex[i]++;
                    }
                }
            }
            for (int i = 0; i < ErrorNumber.Length; i++)
            {
                ErrorMassage[i] = ErrorNumber[i] + "(" + ErrorIndex[i].ToString() + "条)";
            }
            for (int i = 0; i < treeView2.Nodes[0].Nodes.Count; i++)
            {
                treeView2.Nodes[0].Nodes[i].Text = ErrorMassage[i];
            }
        }

        private void btnCopyData_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            //选择保存路径
            string localFilePath = "";
            //string localFilePath, fileNameExt, newFileName, FilePath; 
            SaveFileDialog sfd = new SaveFileDialog();
            //设置文件类型 
            sfd.Filter = "mdb文件(*.mdb)|*.mdb";
            //设置默认文件类型显示顺序 
            sfd.FilterIndex = 1;
            //保存对话框是否记忆上次打开的目录 
            sfd.RestoreDirectory = true;

            sfd.OverwritePrompt = true;
            sfd.FileName = DateTime.Now.ToString("yyyyMMdd") + ".mdb";
            //点了保存按钮进入 
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                localFilePath = sfd.FileName.ToString(); //获得文件路径 
                string fileNameExt = localFilePath.Substring(localFilePath.LastIndexOf("\\") + 1); //获取文件名，不带路径
                string filePath = System.IO.Path.GetDirectoryName(localFilePath);
                IWorkspaceFactory pWorksapceFactory = new AccessWorkspaceFactory();
                IWorkspaceName worksapcename = pWorksapceFactory.Create(filePath, fileNameExt, null, 0);
                IName name = worksapcename as IName;
                IWorkspace pWorkspace = name.Open() as IWorkspace;


                //NewLayersName;
                //ChineseLayerName;
                ArrayList AvaliLayer = new ArrayList();

                for (int i = 0; i < NewLayersName.Length; i++)
                {
                    bool isAvaliable = true;
                    for (int j = 0; j < dgvError.Rows.Count; j++)
                    {
                        if (dgvError.Rows[j].Cells[1].Value.ToString() == NewLayersName[i])
                        {
                            isAvaliable = false;
                            break;
                        }
                    }
                    if (isAvaliable) { AvaliLayer.Add(NewLayersName[i]); }
                }
                for (int i = 0; i < AvaliLayer.Count; i++)
                {
                    IFeatureLayer mCphFeatureLayer = GetLayerByName(AvaliLayer[i].ToString()) as IFeatureLayer;//这是获得要入库的shapefile，获取其FeatureLayer即可
                    //2.创建要素数据集
                    IFeatureClass pCphFeatureClass = mCphFeatureLayer.FeatureClass;
                    //int code = getSpatialReferenceCode(pCphFeatureClass);//参照投影的代号
                    string datasetName = pCphFeatureClass.AliasName;//要素数据集的名称
                    IFeatureDataset pCphDataset = CreateFeatureClass(pWorkspace, pCphFeatureClass, datasetName);
                    //3.导入SHP到要素数据集(
                    importToDB(pCphFeatureClass, pWorkspace, pCphDataset, pCphFeatureClass.AliasName);

                }


            }
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

        private void LineExtract_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            //选择保存路径
            string localFilePath = "";
            //string localFilePath, fileNameExt, newFileName, FilePath; 
            SaveFileDialog sfd = new SaveFileDialog();
            //设置文件类型 
            sfd.Filter = "mdb文件(*.mdb)|*.mdb";
            //设置默认文件类型显示顺序 
            sfd.FilterIndex = 1;
            //保存对话框是否记忆上次打开的目录 
            sfd.RestoreDirectory = true;

            sfd.OverwritePrompt = true;
            sfd.FileName = DateTime.Now.ToString("yyyyMMdd") + "新市镇.mdb";
            //点了保存按钮进入 
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                localFilePath = sfd.FileName.ToString(); //获得文件路径 
                string fileNameExt = localFilePath.Substring(localFilePath.LastIndexOf("\\") + 1); //获取文件名，不带路径
                string filePath = System.IO.Path.GetDirectoryName(localFilePath);
                IWorkspaceFactory pWorksapceFactory = new AccessWorkspaceFactory();
                IWorkspaceName worksapcename = pWorksapceFactory.Create(filePath, fileNameExt, null, 0);
                IName name = worksapcename as IName;
                IWorkspace pWorkspace = name.Open() as IWorkspace;

                IFeatureLayer mCphFeatureLayer = GetLayerByName("YJJBNT") as IFeatureLayer;//这是获得要入库的shapefile，获取其FeatureLayer即可
                //2.创建要素数据集
                if (mCphFeatureLayer != null)
                {
                    IFeatureClass pCphFeatureClass = mCphFeatureLayer.FeatureClass;
                    //int code = getSpatialReferenceCode(pCphFeatureClass);//参照投影的代号
                    string datasetName = pCphFeatureClass.AliasName;//要素数据集的名称
                    IFeatureDataset pCphDataset = CreateFeatureClass(pWorkspace, pCphFeatureClass, datasetName);
                    //3.导入SHP到要素数据集(
                    importToDB(pCphFeatureClass, pWorkspace, pCphDataset, pCphFeatureClass.AliasName);
                }
                mCphFeatureLayer = GetLayerByName("CSKFBJNGHYT") as IFeatureLayer;
                if (mCphFeatureLayer != null)
                {
                    IFeatureClass pCphFeatureClass = mCphFeatureLayer.FeatureClass;
                    //int code = getSpatialReferenceCode(pCphFeatureClass);//参照投影的代号
                    string datasetName = pCphFeatureClass.AliasName;//要素数据集的名称
                    IFeatureDataset pCphDataset = CreateFeatureClass(pWorkspace, pCphFeatureClass, datasetName);
                    //3.导入SHP到要素数据集(
                    importToDB(pCphFeatureClass, pWorkspace, pCphDataset, pCphFeatureClass.AliasName);
                }


                mCphFeatureLayer = GetLayerByName("STKJKZX") as IFeatureLayer;
                IFields pFieldsa = mCphFeatureLayer.FeatureClass.Fields;
                int zdCount = pFieldsa.FieldCount;
                string fileName = "";
                fileName = mCphFeatureLayer.Name;
                ILayer yLayer = GetLayerByName("STKJKZX");
                IFeatureLayer yFeatureLayer = yLayer as IFeatureLayer;          //获取esriGeometryType,作为参数传入新建shp文件的函数来确定新建类型
                IFeatureClass yFeatureClass = yFeatureLayer.FeatureClass;
                string fieldname = yFeatureClass.ShapeFieldName;
                IFields yFields = yFeatureClass.Fields;
                int ind = yFields.FindField(fieldname);
                IField yField = yFields.get_Field(ind);

                //IGeometryDef geometryDef = yField.GeometryDef;
                //esriGeometryType type = geometryDef.GeometryType;   //获取esriGeometryType,作为参数传入新建shp文件的
                //IGeometryDefEdit geoDefEdit = geometryDef as IGeometryDefEdit;
                //geoDefEdit.HasZ_2 = false;

                IFeatureLayer pFeatureLayer = new FeatureLayerClass();          //定义被复制图层和空白shp文件要素和要素类
                IFeatureClass pFeatureClass = pFeatureLayer.FeatureClass;
                pFeatureClass = CT.CreateMemoryFeatureClass(mCphFeatureLayer.FeatureClass);

                ILayer pLayer1 = GetLayerByName("STKJKZX");
                IFeatureLayer pFeatureLayer1 = pLayer1 as IFeatureLayer;
                IFeatureClass pFeatureClass1 = pFeatureLayer1.FeatureClass;

                IQueryFilter pQueryFilter = new QueryFilterClass();             //SQL Filter 
                pQueryFilter.WhereClause = "GKDJ LIKE '01' OR GKDJ LIKE '02'";

                IFeatureCursor pFeatureCursor = pFeatureClass1.Search(pQueryFilter, false);     // 添加要素至要素类并作为一个新涂层展现出来

                IFeature pFeature = pFeatureCursor.NextFeature();
                if (pFeature != null)
                {
                    if (pFeatureClass.Fields.FieldCount != pFeature.Fields.FieldCount)
                    {
                        addFields(pFeature, pFeatureClass, zdCount);
                    }
                    while (pFeature != null)
                    {
                        AddFeatureToFeatureClass(pFeatureClass, pFeature);
                        pFeature = pFeatureCursor.NextFeature();
                    }
                    if (pFeatureClass != null)
                    {
                        IFeatureClass pCphFeatureClass = pFeatureClass;
                        //int code = getSpatialReferenceCode(pCphFeatureClass);//参照投影的代号
                        string datasetName = "STKJKZX";//要素数据集的名称
                        IFeatureDataset pCphDataset = CreateFeatureClass(pWorkspace, pCphFeatureClass, datasetName);
                        //3.导入SHP到要素数据集(
                        importToDB(pCphFeatureClass, pWorkspace, pCphDataset, "STKJKZX");
                    }
                }


                mCphFeatureLayer = GetLayerByName("JSYDKZX") as IFeatureLayer;
                if (mCphFeatureLayer != null)
                {
                    IFeatureClass pCphFeatureClass = mCphFeatureLayer.FeatureClass;
                    //int code = getSpatialReferenceCode(pCphFeatureClass);//参照投影的代号
                    string datasetName = pCphFeatureClass.AliasName;//要素数据集的名称
                    IFeatureDataset pCphDataset = CreateFeatureClass(pWorkspace, pCphFeatureClass, datasetName);
                    //3.导入SHP到要素数据集(
                    importToDB(pCphFeatureClass, pWorkspace, pCphDataset, pCphFeatureClass.AliasName);
                }
                MessageBox.Show("导出成功");
            }
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
                }
                pDT = LD.ShowTableInDataGridView_zenjian(axMapControl1.get_Layer(0) as ITable, dgvTable, out FieldName);
            }
        }
        //导出附表1

        private void buttonCommand1_Click_2(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            if (dgvStastic.IsCurrentCellInEditMode == true)
            {
                dgvStastic.CurrentCell = null;
            }
            string filePath = "";
            SaveFileDialog s = new SaveFileDialog();
            s.Title = "保存Excel文件";
            s.Filter = "Excel文件(*.xlsx)|*.xlsx";
            s.FilterIndex = 1;
            if (s.ShowDialog() == DialogResult.OK)
            {
                filePath = s.FileName;
                string AppendText = "\r\n导出统计表格\r\n时间:" + DateTime.Now.ToString();
                this.Invoke(this.myDelegateAppendTextInfo, new object[] { AppendText });
                if (dgvStastic.Rows.Count <= 0)
                {
                    this.Invoke(this.myDelegateAppendTextInfo, new object[] { "\r\n提示：无数据导出" }); return;
                }
                DataTable tmpStatisticDataTable = new DataTable("StatisticDT");
                DataTable modelTable = new DataTable("ModelTable");
                for (int column = 0; column < dgvStastic.Columns.Count; column++)
                {
                    if (dgvStastic.Columns[column].Visible == true)
                    {
                        DataColumn tempColumn = new DataColumn(dgvStastic.Columns[column].HeaderText, typeof(string));
                        tmpStatisticDataTable.Columns.Add(tempColumn);

                        DataColumn modelColumn = new DataColumn(dgvStastic.Columns[column].Name, typeof(string));
                        modelTable.Columns.Add(modelColumn);
                    }
                }
                for (int row = 0; row < dgvStastic.Rows.Count; row++)
                {
                    DataRow tempRow = tmpStatisticDataTable.NewRow();
                    for (int i = 0; i < tmpStatisticDataTable.Columns.Count; i++)
                    {
                        tempRow[i] = dgvStastic.Rows[row].Cells[modelTable.Columns[i].ColumnName].Value;
                    }
                    tmpStatisticDataTable.Rows.Add(tempRow);
                }
                if (tmpStatisticDataTable == null)
                {
                    return;
                }
                //第二步：导出dataTable到Excel  
                long rowNum = tmpStatisticDataTable.Rows.Count;//行数  
                int columnNum = tmpStatisticDataTable.Columns.Count;//列数  
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
                        datas[0, i] = tmpStatisticDataTable.Columns[i].Caption;
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
                            object obj = tmpStatisticDataTable.Rows[r][tmpStatisticDataTable.Columns[i].ToString()];
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
    }
}