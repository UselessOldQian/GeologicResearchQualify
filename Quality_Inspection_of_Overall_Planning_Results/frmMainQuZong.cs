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

namespace Quality_Inspection_of_Overall_Planning_Results
{
    public partial class frmMainQuZong : Form
    {

        private frmLoading loadForm;
        MapAction _mapAction = MapAction.Null;
        IFeatureWorkspace pFeatureWorkspace;
        IFeatureLayer pFeatureLayer;
        IFeatureDataset pFeatureDataset;
        ILayer selectedLayer;
        static DataTable pDT;

        public delegate void AppendTextInfo(string strMsg);
        public AppendTextInfo myDelegateAppendTextInfo;

        public delegate void UpdateBarValue(int iValue);
        public UpdateBarValue myDelegateUpdateBarValue;

        public delegate void UpdateUiStatus(string strMsg);
        public UpdateUiStatus myDelegateUpdateUiStatus;

        public int ProcessBarMaxValue = 0;
        public bool IsRun = false;
        List<String> FieldName = new List<string>();

        public frmMainQuZong()
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
                case "JCDLYS":
                    return EnglishName + "(基础地理要素)";
                case "JQDLTB":
                    return EnglishName + "(基期地类图斑)";
                case "STBHKZX":
                    return EnglishName + "(生态保护控制线)";
                case "YJJBNTBHQ":
                    return EnglishName + "(基本农田保护区)"; 
                case "CSKFBJ":
                    return EnglishName + "(城市开发边界)";
                case "TDYTQ2035":
                    return EnglishName + "(土地用途区)";
                case "JSYDHJBNTGZ2035":
                    return EnglishName + "(建设用地和基本农田管制2035)";
                case "TDZZ":
                    return EnglishName + "(土地整治)";
                case "ZBTZQKB":
                    return EnglishName + "(指标条情况表)";
                case "ZBFJB":
                    return EnglishName + "(指标分解表)";
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
                    axMapControl1.ActiveView.FocusMap.get_Layer(0).Visible = false;
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
            pDT = LD.ShowTableInDataGridView(ptable, dgvTable, ref pCursor, ref pRrow,out FieldName);
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
                pDT = LD.ShowTableInDataGridView((ITable)selectedLayer, dgvTable, ref pCursor, ref pRrow,out FieldName);
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

        static ICursor pCursor;
        static IRow pRrow;
        private void dgvTable_Scroll(object sender, ScrollEventArgs e)
        {
            if (e.ScrollOrientation == ScrollOrientation.VerticalScroll &&
                (e.NewValue + dgvTable.DisplayedRowCount(false) == dgvTable.Rows.Count))
            {
                this.Invoke(myDelegateUpdateUiStatus, new object[] { "正在加载数据..." });
                DataTable tablet = LD.GetData(ref pCursor, ref pRrow);
                if (tablet != null)
                {
                    pDT.Merge(tablet);//表合并
                    int r = Convert.ToInt16(e.NewValue);//保存当前滚动条的位置
                    this.dgvTable.DataSource = pDT;
                    dgvTable.FirstDisplayedScrollingRowIndex = r;//滚动条回到触发滚动事件时的位置                 
                }
            }
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "数据加载完成" });
        }

        DataTable pErrorDataTable = new DataTable();
        public string[] NewLayersName = { "JCDLYS", "JQDLTB","STBHKZX", "YJJBNTBHQ", "CSKFBJ", "TDYTQ2035", "JSYDHJBNTGZ2035", "TDZZ"};
        public string[] ChineseLayerName = { "基础地理要素", "基础地类图斑", "生态保护控制线", "基本农田保护区", "城市开发边界", "土地用途区", "建设用地和基本农田管制", "土地整治" };
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
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR1101:" + GetChineseName(NewLayersName[i]) +"不存在" });
                    pErrorDataTable.Rows.Add(new object[] { "1101", NewLayersName[i], null, null, GetChineseName(NewLayersName[i]) + "不存在", false, true });
                    continue;
                }
                IFeatureClass pFeaCls = (layerresult as IFeatureLayer).FeatureClass;
                //再通过IGeoDataset接口获取FeatureClass坐标系统
                ISpatialReference pSpatialRef = (pFeaCls as IGeoDataset).SpatialReference;
                if (pSpatialRef.Name.ToUpper() == "UNKNOWN")
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR2201:" + GetChineseName(NewLayersName[i]) + "投影为" + pSpatialRef.Name });
                    pErrorDataTable.Rows.Add(new object[] { "2201", NewLayersName[i], null, null, GetChineseName(NewLayersName[i])+"投影为" + pSpatialRef.Name,false, true });
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
                case "JCDLYS":
                    return "基础地理要素";
                case "JQDLTB":
                    return "基础地类图斑";
                case "STBHKZX":
                    return "生态保护控制线";
                case "YJJBNTBHQ":
                    return "基本农田保护区";
                case "CSKFBJ":
                    return "城市开发边界";
                case "TDYTQ2035":
                    return "土地用途区";
                case "JSYDHJBNTGZ2035":
                    return "建设用地和基本农田管制";
                case "TDZZ":
                    return "土地整治";
                case "ZBTZQKB":
                    return "指标调整情况表";
                case "ZBFJB":
                    return "指标分解表";
                case "TDLYJGTZB":
                    return "土地利用结构调整表";
                case "GDZBPHB":
                    return "耕地占补平衡表";
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
                string Chinesename=switchName(strLayerName);
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
            ILayer layerresult = GetLayerByName("JCDLYS");
            if (layerresult != null)
            {
                CheckAttributeBSM(layerresult);
                CheckTextAttribute(layerresult, "XZBJMC", null, 20);
                CheckAttributeMJ(layerresult, "XZBJMJ");
                CheckAttributeMSorSM(layerresult, "SM");
                this.Invoke(myDelegateUpdateUiStatus, new object[] { "图层JCDLYS属性检查完成" });
            }
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });

            layerresult = GetLayerByName("JQDLTB");
            if (layerresult != null)
            {
                CheckAttributeBSM(layerresult);
                CheckAttributeDLBM_SX(layerresult);
                CheckTextAttribute(layerresult, "GHJSFLMC", null, 20);
                CheckTextAttribute(layerresult, "GHJSFLBM", null, 5);
                CheckAttributeMJ(layerresult, "TBMJ");
                this.Invoke(myDelegateUpdateUiStatus, new object[] { "图层JQDLTB属性检查完成" });
            }
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });

            layerresult = GetLayerByName("STBHKZX");
            if (layerresult != null)
            {
                CheckAttributeBSM(layerresult);
                CheckTextAttribute(layerresult, "GKDJDM", GKDJDMrange, 10);
                CheckTextAttribute(layerresult, "GKDJMC", GKDJMCrange, 20);
                CheckAttributeMJ(layerresult, "MJ");
                CheckAttributeMSorSM(layerresult, "SM");
                this.Invoke(myDelegateUpdateUiStatus, new object[] { "图层STBHKZX属性检查完成" });
            }
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });

            layerresult = GetLayerByName("YJJBNTBHQ");
            if (layerresult != null)
            {
                CheckAttributeBSM(layerresult);
                CheckTextAttribute(layerresult, "BHQBH", null, 4);
                CheckDoubleAttribute(layerresult, "BHQMJ");
                CheckDoubleAttribute(layerresult, "NYDMJ");
                CheckDoubleAttribute(layerresult, "GDMJ");
                CheckAttributeMJ(layerresult, "JBNTMJ");
                CheckAttributeMSorSM(layerresult, "SM");
                this.Invoke(myDelegateUpdateUiStatus, new object[] { "图层YJJBNTBHQ属性检查完成" });
            }
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });

            layerresult = GetLayerByName("CSKFBJ");
            if (layerresult != null)
            {
                CheckAttributeBSM(layerresult);
                CheckTextAttribute(layerresult, "XZQDM", null, 20);
                CheckTextAttribute(layerresult, "XZQMC", null, 20);
                CheckAttributeMJ(layerresult, "QYMJ");
                CheckAttributeMSorSM(layerresult, "SM");
                this.Invoke(myDelegateUpdateUiStatus, new object[] { "图层CSKFBJ属性检查完成" });
            }
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });

            layerresult = GetLayerByName("TDYTQ2035");
            if (layerresult != null)
            {
                CheckAttributeBSM(layerresult);
                CheckTextAttribute(layerresult, "XZQDM", null, 20);
                CheckTextAttribute(layerresult, "XZQMC", null, 20);
                CheckTextAttribute(layerresult, "TDYTQLXDM", TDYTQLXDMrange, 10);
                CheckTextAttribute(layerresult, "TDYTQLXMC", null, 10);
                CheckAttributeMJ(layerresult, "TDYTQMJ");
                CheckAttributeMSorSM(layerresult, "SM");
                this.Invoke(myDelegateUpdateUiStatus, new object[] { "图层TDYTQ2035属性检查完成" });
            }
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });

            layerresult = GetLayerByName("JSYDHJBNTGZ2035");
            if (layerresult != null)
            {
                CheckAttributeBSM(layerresult);
                CheckAttributeXZQDM(layerresult);
                CheckAttributeXZQMC(layerresult);
                CheckTextAttribute(layerresult, "GZQLXDM", GZQLXDMrange, 3);
                CheckTextAttribute(layerresult, "GZQLXMC", null, 20);
                CheckAttributeMJ(layerresult, "GZQMJ");
                this.Invoke(myDelegateUpdateUiStatus, new object[] { "图层JSYDHJBNTGZ2035属性检查完成" });
            }
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });

            layerresult = GetLayerByName("TDZZ");
            if (layerresult != null)
            {
                CheckAttributeBSM(layerresult);
                CheckTextAttribute(layerresult, "TDZZLX", TDZZLXrange, 10);
                CheckAttributeMJ(layerresult, "MJ");
                this.Invoke(myDelegateUpdateUiStatus, new object[] { "图层TDZZ属性检查完成" });
            }
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateUpdateBarValue, new object[] { 0 });
            BindingSource bind = new BindingSource();
            bind.DataSource = pErrorDataTable;
            dgvError.DataSource = bind;
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "属性检查完成" });
            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\n属性检查完成\r\n" });
        }

        private void CheckDoubleAttribute(ILayer player,string DoubleField)
        {
            int attr = (player as ITable).FindField(DoubleField);
            if (attr < 0)
            {
                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + player.Name + "的属性字段" + DoubleField + "不存在或正名命名错误" });
                pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, DoubleField, null, player.Name + "的属性字段" + DoubleField + "不存在或正名命名错误", false, true });
            }
            IField pfield = (player as ITable).Fields.get_Field(attr);
            if (esriFieldType.esriFieldTypeDouble != pfield.Type)
            {
                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段" + DoubleField + "类型不是Double" });
                pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, DoubleField, null, GetChineseName(player.Name) + "的属性字段" + DoubleField + "类型不是Double", false, true });
            }
            ICursor pCursor = (player as ITable).Search(null, false);
            IRow pRrow = pCursor.NextRow();
            while (pRrow != null)
            {
                if (Convert.IsDBNull(pRrow.get_Value(attr)))
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3601:" + GetChineseName(player.Name) + "的属性字段" + DoubleField + " objectID=" + pRrow.OID + "的值为空" });
                    pErrorDataTable.Rows.Add(new object[] { "3601", player.Name, DoubleField, pRrow.OID, GetChineseName(player.Name) + "的属性字段" + DoubleField + " objectID=" + pRrow.OID + "的值为空", false, true });
                }
                else if (Double.Parse(pRrow.get_Value(attr).ToString()) <= 0)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3401:" + GetChineseName(player.Name) + "的属性字段" + DoubleField + " objectID=" + pRrow.OID + "的值不在值域内" });
                    pErrorDataTable.Rows.Add(new object[] { "3401", player.Name, DoubleField, pRrow.OID.ToString(), GetChineseName(player.Name) + "的属性字段" + DoubleField + " objectID=" + pRrow.OID + "的值不在值域内", false, true });
                }
                pRrow = pCursor.NextRow();
            }
 
        }


        private void CheckGHFWAttribute()
        {
            ILayer player = GetLayerByName("GHFW");
            if (player == null){return;}
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
                                YSDMValue="2003010100";
                                break;
                            case "CSKFBJNGHYT":
                            case "城市开发边界内规划用途":
                                YSDMValue="2003020241";
                                break;
                            case "JSYDKZX":
                            case "建设用地控制线":
                                YSDMValue="2003020140";
                                break;
                            case "STKJKZX":
                            case "生态空间控制线":
                                YSDMValue="2003020120";
                                break;
                            case "YJJBNT":
                            case "永久基本农田":
                                YSDMValue="2003020110";
                                break;
                            case "JSYDHJBNTGZ2035":
                            case "建设用地和基本农田管制":
                                YSDMValue="2003020221";
                                break;
                            case "JLHDK":
                            case "减量化地块":
                                YSDMValue="2003020510";
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
                if (pfield.Length != 20)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段XZQDM字段长度不为20" });
                    pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "XZQDM", null, GetChineseName(player.Name) + "的属性字段XZQDM字段长度不为20", false, true });
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
                if (pfield.Length != 20)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段XZQMC字段长度不为20" });
                    pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, "XZQMC", null, GetChineseName(player.Name) + "的属性字段XZQMC字段长度不为20", false, true });
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
                        this.Invoke(myDelegateAppendTextInfo,new object[] {"\r\nERROR3601:" + player.Name + "的属性字段DLBM_SX objectID=" + pRrow.get_Value(IDIndex) + "的值为空"}); 
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

        string[] GHYTrange = { "城镇建设用地区", "产业基地", "产业社区", "战略预留区", "规划水域","010","021","022","030","040" };
        string[] LXrange = { "城市开发边界内建设用地", "其他建设用地区" };

        /// <summary>
        /// 检查GHYT字段
        /// </summary>
        /// <param name="player">要检查图层的名称</param>
        /// <param name="GHYTorLX">字段名称为GHYT还是LX</param>
        private void CheckAttributeGHYT(ILayer player, string GHYTorLX)
        {
            string[] valuerange = { };
            int textrange = 0;
            if (GHYTorLX == "GHYT") { valuerange = GHYTrange; textrange = 10; }
            if (GHYTorLX == "LX") { valuerange = LXrange; textrange = 20; }
            ITable ptable = (ITable)player;
            int FieldIndex = ptable.FindField(GHYTorLX);
            if (FieldIndex < 0)
            {
                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段" + GHYTorLX + "不存在或正名命名错误" });
                pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, GHYTorLX, null, GetChineseName(player.Name) + "的属性字段" + GHYTorLX + "不存在或正名命名错误", false, true });
                return;
            }
            int IDIndex = ptable.FindField("OBJECTID");
            IField pfield = ptable.Fields.get_Field(FieldIndex);
            if (pfield != null)
            {
                if (esriFieldType.esriFieldTypeString != pfield.Type)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段GHYT类型不是Text" });
                    pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, GHYTorLX, null, GetChineseName(player.Name) + "的属性字段GHYT类型不是Text", false, true });
                    return;
                }
                if (pfield.Length != textrange)
                {
                    this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段GHYT字段长度不为10" });
                    pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, GHYTorLX, null, GetChineseName(player.Name) + "的属性字段GHYT字段长度不为10", false, true });
                }
                ICursor pCursor = ptable.Search(null, false);
                IRow pRrow = pCursor.NextRow();
                while (pRrow != null)
                {
                    if (Convert.IsDBNull(pRrow.get_Value(FieldIndex)))
                    {
                        this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3601:" + GetChineseName(player.Name) + "的属性字段GHYT objectID=" + pRrow.get_Value(IDIndex) + "的值为空" });
                        pErrorDataTable.Rows.Add(new object[] { "3601", player.Name, GHYTorLX, pRrow.get_Value(IDIndex).ToString(), GetChineseName(player.Name) + "的属性字段GHYT objectID=" + pRrow.get_Value(IDIndex) + "的值为空", false, true });
                    }
                    else
                    {
                        if (valuerange.Contains((string)pRrow.get_Value(FieldIndex)) == false)
                        {
                            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3301:" + GetChineseName(player.Name) + "的属性字段GHYT objectID=" + pRrow.get_Value(IDIndex) + "的值不符合要求" });
                            pErrorDataTable.Rows.Add(new object[] { "3301", player.Name, GHYTorLX, pRrow.get_Value(IDIndex).ToString(), GetChineseName(player.Name) + "的属性字段GHYT objectID=" + pRrow.get_Value(IDIndex) + "的值不符合要求", false, true });
                        }
                    }
                    pRrow = pCursor.NextRow();
                }
            }
            else
            {
                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3201:" + GetChineseName(player.Name) + "的属性字段GHYT不存在或正名命名错误" });
                pErrorDataTable.Rows.Add(new object[] { "3201", player.Name, GHYTorLX, pRrow.get_Value(IDIndex).ToString(), GetChineseName(player.Name) + "的属性字段GHYT不存在或正名命名错误", false, true });
                return;
            }
        }

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
        string[] GKDJDMrange = { "01", "02", "03", "04" };
        string[] GKDJMCrange = { "一类生态空间", "二类生态空间", "三类生态空间", "四类生态空间" };
        string[] JQDLDMrange = { "20", "22", "25", "26", "27" };
        string[] JQDLMCrange = { "城镇建设用地", "工业仓储用地", "农村居民点用地", "交通运输用地", "其他建设用地" };
        string[] SSSXrange = { "近期", "远期" };
        string[] TDYTQLXDMrange = { "110", "1101", "1102", "1103", "1104", "1105", "1106", "1107", "1108", "1109", "1110", "1111", "1112", "010","020","120","210" };
        string[] TDZZLXrange = { "01", "02", "03" };
        string[] GZQLXDMrange = { "01", "011", "012", "02", "021", "022", "03", "04" };
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
                        if (TextRange != null)
                        {
                            if (TextRange.Contains((string)pRrow.get_Value(FieldIndex)) == false)
                            {
                                this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\nERROR3301:" + GetChineseName(player.Name) + "的属性字段" + TextAttributeName + " objectID=" + pRrow.get_Value(IDIndex) + "的值不符合要求" });
                                pErrorDataTable.Rows.Add(new object[] { "3301", player.Name, TextAttributeName, pRrow.get_Value(IDIndex).ToString(), GetChineseName(player.Name) + "的属性字段" + TextAttributeName + " objectID=" + pRrow.get_Value(IDIndex) + "的值不符合要求", false, true });
                            }
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
            progressBar1.Maximum = 2;
            ILayer layer1 = GetLayerByName("JCDLYS");
            ILayer layer2 = GetLayerByName("JQDLTB");
            ILayer layer3 = GetLayerByName("TDYTQ2035");
            ILayer layer4 = GetLayerByName("JSYDHJBNTGZ2035");
            ILayer[] Layerlist = { layer1, layer2, layer3,layer4 };
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            //this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeConsistent1(Layerlist, ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            //this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeConsistent2(GetLayerByName("行政区界") as IFeatureLayer, layer1 as IFeatureLayer, ref pErrorDataTable, "行政区界", "6401") });
            //this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            //this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("JBNTBHTB") as IFeatureLayer, GetLayerByName("YJJBNT") as IFeatureLayer, ref pErrorDataTable, "行政区界", "6401") });
            //this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            //this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("生态保护红线") as IFeatureLayer, GetLayerByName("STKJKZX") as IFeatureLayer, ref pErrorDataTable, "行政区界", "BHLX LIKE '11*'", "6401") });
            //this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            //this.Invoke(myDelegateAppendTextInfo, new object[] { CheckSpatial6401_5(GetLayerByName("CSKFBJNGHYT") as IFeatureLayer, GetLayerByName("城市开发边界") as IFeatureLayer, ref pErrorDataTable) });            
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
            IGeometry pDiff3 = pGeoInTP3.Intersect(XZQ_geo,esriGeometryDimension.esriGeometry2Dimension);
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
            string JCDLYS = "JCDLYS";
            string JQDLTB = "JQDLTB";
            string STBHKZX = "STBHKZX";
            string YJJBNTBHQ = "YJJBNTBHQ";
            string CSKFBJ = "CSKFBJ";
            string TDYTQ2035 = "TDYTQ2035";
            string JSYDHJBNTGZ2035 = "JSYDHJBNTGZ2035";
            string TDZZ = "TDZZ";
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfIntersection(GetLayerByName(JCDLYS), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfIntersection(GetLayerByName(JQDLTB), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfIntersection(GetLayerByName(STBHKZX), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfIntersection(GetLayerByName(YJJBNTBHQ), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfIntersection(GetLayerByName(CSKFBJ), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfIntersection(GetLayerByName(TDYTQ2035), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfIntersection(GetLayerByName(JSYDHJBNTGZ2035), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfIntersection(GetLayerByName(TDZZ), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSimple(GetLayerByName(JCDLYS), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSimple(GetLayerByName(JQDLTB), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSimple(GetLayerByName(STBHKZX), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSimple(GetLayerByName(YJJBNTBHQ), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSimple(GetLayerByName(CSKFBJ), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSimple(GetLayerByName(TDYTQ2035), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSimple(GetLayerByName(JSYDHJBNTGZ2035), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSimple(GetLayerByName(TDZZ), ref pErrorDataTable) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfOverlap(GetLayerByName(JCDLYS), ref pErrorDataTable, 1) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfOverlap(GetLayerByName(JQDLTB), ref pErrorDataTable, 10000) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfOverlap(GetLayerByName(STBHKZX), ref pErrorDataTable, 100) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfOverlap(GetLayerByName(YJJBNTBHQ), ref pErrorDataTable, 100) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfOverlap(GetLayerByName(CSKFBJ), ref pErrorDataTable, 100) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfOverlap(GetLayerByName(TDYTQ2035), ref pErrorDataTable, 1000) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfOverlap(GetLayerByName(JSYDHJBNTGZ2035), ref pErrorDataTable, 100) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CT.CheckSelfOverlap(GetLayerByName(TDZZ), ref pErrorDataTable, 100) });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            Check4301(GetLayerByName("JCDLYS"), "XZBJMJ");
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            Check4301(GetLayerByName("JQDLTB"), "TBMJ");
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            Check4301(GetLayerByName("STBHKZX"), "MJ");
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            Check4301(GetLayerByName("YJJBNTBHQ"), "BHQMJ");
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            Check4301(GetLayerByName("CSKFBJ"), "QYMJ");
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            Check4301(GetLayerByName("TDYTQ2035"), "TDYTQMJ");
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            Check4301(GetLayerByName("JSYDHJBNTGZ2035"), "GZQMJ");
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            Check4301(GetLayerByName("TDZZ"), "MJ");
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
            if (AreaCode.ComboBox.SelectedItem == null)
            {
                MessageBox.Show("请选择要检查的区域");
                return;
            }
            else if (AreaSwitchCode(AreaCode.ComboBox.SelectedItem.Value.ToString()) == null)
            {
                MessageBox.Show("请选择要检查的区域");
                return;
            }
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "正在规划布局检查..." });
            if (cbIsClear.Checked == true)
            {
                pErrorDataTable.Rows.Clear();
            }
            progressBar1.Maximum = 6;
            string AppendText = "\r\n规划布局检查\r\n时间:" + DateTime.Now.ToString();
            string condition = "JBNTTBBH LIKE " + AreaSwitchCode(AreaCode.ComboBox.SelectedItem.Value.ToString());
            this.Invoke(this.myDelegateAppendTextInfo, new object[] { AppendText });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals_2035_6501(GetLayerByName("JBNTBHTB") as IFeatureLayer,GetLayerByName("JSYDHJBNTGZ2035") as IFeatureLayer,ref pErrorDataTable,
                "JBNTBHTB",condition,null,"GZQLXDM LIKE '03'","6501"  )});
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("TDYTQ2035") as IFeatureLayer, GetLayerByName("YJJBNTBHQ") as IFeatureLayer, ref pErrorDataTable, "TDYTQ2035", "TDYTQLXDM LIKE '010'", "YJJBNTBHQ", null, "RELATE(G1, G2, 'T*F*T*F**')", "6503") });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("TDYTQ2035") as IFeatureLayer, GetLayerByName("CSKFBJ") as IFeatureLayer, ref pErrorDataTable, "TDYTQ2035", "TDYTQLXDM LIKE '11*' OR TDYTQLXDM LIKE '1106'", "CSKFBJ", null, "RELATE(G1, G2, 'T*F*T*F**')", "6503") });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("TDYTQ2035") as IFeatureLayer, GetLayerByName("JSYDHJBNTGZ2035") as IFeatureLayer, ref pErrorDataTable, "TDYTQ2035", "TDYTQLXDM LIKE '11*' AND TDYTQLXDM <> '1106'", "JSYDHJBNTGZ2035", "GZQLXDM LIKE '01*' OR GZQLXDM LIKE '02*'", "RELATE(G1, G2, 'T*F*T*F**')", "6504") });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("TDYTQ2035") as IFeatureLayer, GetLayerByName("STBHKZX") as IFeatureLayer, ref pErrorDataTable, "TDYTQ2035", "TDYTQLXDM LIKE '210'", "STBHKZX", "GKDJDM LIKE '01*' OR GKDJDM LIKE '02*'", "RELATE(G1, G2, 'T*F*T*F**')", "6504") });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("TDYTQ2035") as IFeatureLayer, GetLayerByName("JSYDHJBNTGZ2035") as IFeatureLayer, ref pErrorDataTable, "TDYTQ2035", "TDYTQLXDM LIKE '210'", "JSYDHJBNTGZ2035", "GZQLXDM LIKE '04'", "RELATE(G1, G2, 'T*F*T*F**')", "6504") });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("STBHKZX") as IFeatureLayer, GetLayerByName("JSYDHJBNTGZ2035") as IFeatureLayer, ref pErrorDataTable, "STBHKZX", "GKDJDM LIKE '01*' OR GKDJDM LIKE '02*'", "JSYDHJBNTGZ2035", "GZQLXDM LIKE '04'", "RELATE(G1, G2, 'T*F*T*F**')", "6504") });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("TDYTQ2035") as IFeatureLayer, GetLayerByName("生态保护红线") as IFeatureLayer, ref pErrorDataTable, "TDYTQ2035", "TDYTQLXDM LIKE '210'", "生态保护红线", condition, "RELATE(G1, G2, 'T*F*T*F**')", "6504") });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("STBHKZX") as IFeatureLayer, GetLayerByName("生态保护红线") as IFeatureLayer, ref pErrorDataTable, "STBHKZX", "GKDJDM LIKE '01*' OR GKDJDM LIKE '02*'", "生态保护红线", condition, "RELATE(G1, G2, 'T*F*T*F**')", "6504") });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals(GetLayerByName("JSYDHJBNTGZ2035") as IFeatureLayer, GetLayerByName("生态保护红线") as IFeatureLayer, ref pErrorDataTable, "JSYDHJBNTGZ2035", "GZQLXDM LIKE '04'", "生态保护红线", condition, "RELATE(G1, G2, 'T*F*T*F**')", "6504") });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckSpatialRangeEquals_2035_6501(GetLayerByName("TDYTQ2035") as IFeatureLayer,GetLayerByName("JSYDHJBNTGZ2035") as IFeatureLayer,ref pErrorDataTable,
                "TDYTQ2035","TDYTQLXDM LIKE '1106'","JSYDHJBNTGZ2035","GZQLXDM LIKE '03'","6504"  )});
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
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "正在落图的允许建设区规模统计..." });
            progressBar1.Maximum = 2;
            string AppendText = "\r\n落图的允许建设区规模统计\r\n时间:" + DateTime.Now.ToString();
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            StatisticTable = CDC.getLayerAreaByCity_1(GetLayerByName("JSYDHJBNTGZ2035"), "GZQLXDM LIKE '01*' OR GZQLXDM LIKE '02*'", GetLayerByName("JCDLYS"));
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            BindingSource bind = new BindingSource();
            bind.DataSource = StatisticTable;
            dgvStastic.DataSource = bind;
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "落图的允许建设区规模统计完毕" });
            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\n落图的允许建设区规模统计完成\r\n完成时间:" + DateTime.Now.ToString() + "\r\n" });
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
            axMapControl1.AddShapeFile(pFolder,pFileName);
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
            this.buttonCommand2_Click(sender, e);
            this.buttonCommand3_Click(sender, e);
            this.buttonCommand5_Click(sender, e);
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "所有检查完毕" });
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "所有检查完毕" });
            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\n所有检查完成\r\n完成时间:" + DateTime.Now.ToString() + "\r\n***************************************************************************************************************************************************************************\r\n" });
            this.Invoke(myDelegateUpdateBarValue, new object[] { 0 });
            cbIsClear.Checked = true;
        }

        private void dgvError_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {

            string ColumnName = this.dgvError.Columns[e.ColumnIndex].Name;
            if (ColumnName == "ErrorCheck"||ColumnName=="ErrorExcept")
            {
                if (dgvError.IsCurrentCellInEditMode == true) 
                { 
                    dgvError.CurrentCell = null; 
                }
                if(Convert.IsDBNull(dgvError.Rows[0].Cells[e.ColumnIndex].Value))
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
            for(int i=0;i<dgvError.Rows.Count;i++)
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
            e.Cancel=true;
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
                else if(pDataset is ITable)
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
                ILayer player=GetLayerByName(dgvError.CurrentRow.Cells[1].Value.ToString());
                IFeatureSelection pFeatureSelection = (player as IFeatureLayer) as IFeatureSelection;
                //创建过滤器
                IQueryFilter pQueryFilter = new QueryFilterClass();
                //设置过滤器对象的查询条件
                pQueryFilter.WhereClause = "OBJECTID = "+OID.ToString();
                //根据查询条件选择要素
                pFeatureSelection.SelectFeatures(pQueryFilter, esriSelectionResultEnum.esriSelectionResultNew, false);
                ISimpleFillSymbol SFS=new SimpleFillSymbolClass();
                ISimpleLineSymbol ILS=new SimpleLineSymbolClass();
                SFS.Style = esriSimpleFillStyle.esriSFSSolid;
                SFS.Color = getRGB(255, 0, 0);
                ILS.Color = getRGB(0, 255, 0);
                ILS.Style = esriSimpleLineStyle.esriSLSSolid;
                ILS.Width = 13;
                SFS.Outline=ILS;
                pFeatureSelection.SelectionSymbol = SFS as ISymbol;
                IArea pArea = (player as IFeatureLayer).FeatureClass.GetFeature(OID).Shape as IArea;
                IPoint iPnt = pArea.LabelPoint;
                axMapControl1.Extent = (player as IFeatureLayer).FeatureClass.GetFeature(OID).Shape.Envelope;
                axMapControl1.CenterAt(iPnt);
                ShowAllTableInDataGridView(player as ITable, dgvTable, OID);
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

        public void ShowAllTableInDataGridView(ITable ptable, DataGridView DGV,int OID)
        {
            DGV.DataSource = null;
            DataTable pDataTable = new DataTable();//建立一个table
            for (int i = 0; i < ptable.Fields.FieldCount; i++)
            {
                string FieldName;//建立一个string变量存储Field的名字
                FieldName = ptable.Fields.get_Field(i).AliasName;
                pDataTable.Columns.Add(FieldName);
            }
            int index = 0;
            int rowindex = 0;
            pCursor = ptable.Search(null, false);
            pRrow = pCursor.NextRow();
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
                if(OID==pRrow.OID){rowindex = index;}
                pRrow = pCursor.NextRow();
                index++;
            }
            DGV.DataSource = pDataTable;
            DGV.Rows[rowindex].Selected = true;
            DGV.FirstDisplayedScrollingRowIndex = rowindex;
        }

        private void btnTextAuxiliaryCheck_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            this.StatisticTable.Rows.Clear();
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "正在城市开发边界情况统计..." });
            progressBar1.Maximum = 2;
            string AppendText = "\r\n城市开发边界情况统计\r\n时间:" + DateTime.Now.ToString();
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            StatisticTable = CDC.getLayerAreaByCity_2(GetLayerByName("TDYTQ2035"), "TDYTQLXDM LIKE '11*'","TDYTQLXDM LIKE '11*' AND TDYTQLXDM <> '1106'", GetLayerByName("JCDLYS"));
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            BindingSource bind = new BindingSource();
            bind.DataSource = StatisticTable;
            dgvStastic.DataSource = bind;
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "城市开发边界情况统计完毕" });
            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\n城市开发边界情况统计完成\r\n完成时间:" + DateTime.Now.ToString() + "\r\n" });
            this.uiTab1.SelectedTab = this.uiTab1.TabPages[2];
            this.Invoke(myDelegateUpdateBarValue, new object[] { 0 });
        }

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
            string[] ErrorNumber=new string[] {"1101","2201","3201","3301","3401","3601","4301","4101","6401","6402","6403","6501","6502","6503","6504","6505"};
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
                string fileNameExt =localFilePath.Substring(localFilePath.LastIndexOf("\\") + 1); //获取文件名，不带路径
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

        public string AreaSwitchCode(string AreaName)
        {
            switch(AreaName)
            {
                case "宝山区":
                    return "'310113*'";
                case "奉贤区":
                    return "310120*";
                case "青浦区":
                    return "310118*";
                case "松江区":
                    return "310117*";
                case "崇明区":
                    return "310230*";
                case "嘉定区":
                    return "310114*";
                case "浦东新区":
                    return "310115*";
                case "金山区":
                    return "310116*";
                case "闵行区":
                    return "310112*";
                default:
                    return null;
            }
        }

        private void buttonCommand2_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            if (AreaCode.ComboBox.SelectedItem == null)
            {
                MessageBox.Show("请选择要检查的区域");
                return;
            }
            else if (AreaSwitchCode(AreaCode.ComboBox.SelectedItem.Value.ToString()) == null)
            {
                MessageBox.Show("请选择要检查的区域");
                return;
            }
            if (cbIsClear.Checked == true)
            {
                pErrorDataTable.Rows.Clear();
            }
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "正在空间与非空间数据一致性检查..." });
            string AppendText = "\r\n空间与非空间数据一致性检查\r\n时间:" + DateTime.Now.ToString();
            this.Invoke(this.myDelegateAppendTextInfo, new object[] { AppendText });
            progressBar1.Maximum = 2;
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckArea(GetLayerByName("JSYDHJBNTGZ2035"), "GZQLXDM LIKE '01*' OR GZQLXDM LIKE '02*'", treeView1, "ZBTZQKB", "ZBDM = '06'", "ZBMJ", 0, ref pErrorDataTable, "6402") });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            string condition = "JBNTTBBH LIKE " + AreaSwitchCode(AreaCode.ComboBox.SelectedItem.Value.ToString());
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckArea(GetLayerByName("JBNTBHTB"), condition, treeView1, "ZBTZQKB", "ZBDM = '02'", "ZBMJ", 1, ref pErrorDataTable, "6402") });
            BindingSource bind = new BindingSource();
            bind.DataSource = pErrorDataTable;
            dgvError.DataSource = bind;
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "空间与非空间数据一致性检查完毕" });
            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\n空间与非空间数据一致性检查完成\r\n完成时间:" + DateTime.Now.ToString() + "\r\n" });
            this.Invoke(myDelegateUpdateBarValue, new object[] { 0 });
        }

        private void buttonCommand3_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            if (AreaCode.ComboBox.SelectedItem == null)
            {
                MessageBox.Show("请选择要检查的区域");
                return;
            }
            else if (AreaSwitchCode(AreaCode.ComboBox.SelectedItem.Value.ToString()) == null)
            {
                MessageBox.Show("请选择要检查的区域");
                return;
            }
            if (cbIsClear.Checked == true)
            {
                pErrorDataTable.Rows.Clear();
            }
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "正在非空间数据一致性检查..." });
            string AppendText = "\r\n非空间数据一致性检查\r\n时间:" + DateTime.Now.ToString();
            string condition = "XZQDM LIKE " + AreaSwitchCode(AreaCode.ComboBox.SelectedItem.Value.ToString());
            this.Invoke(this.myDelegateAppendTextInfo, new object[] { AppendText });
            progressBar1.Maximum = 5;
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            if (listBox1.Items.Count == 0)
            {
                MessageBox.Show("未加入全市指标表");
            }
            else
            {
                this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckTableAndExcelArea("ZBTZQKB", "ZBDM = '01'", "ZBMJ", treeView1, listBox1.Items[0].ToString(), "行政区 LIKE '" + AreaCode.ComboBox.SelectedItem.Value.ToString() + "'", 2, ref pErrorDataTable, "6403", 1, "耕地保有量（万亩）") });
                this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckTableAndExcelArea("ZBTZQKB", "ZBDM = '02'", "ZBMJ", treeView1, listBox1.Items[0].ToString(), "行政区 LIKE '" + AreaCode.ComboBox.SelectedItem.Value.ToString() + "'", 2, ref pErrorDataTable, "6403", 2, "基本农田保护面积（万亩）") });
                this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckTableAndExcelArea("ZBTZQKB", "ZBDM = '06'", "ZBMJ", treeView1, listBox1.Items[0].ToString(), "行政区 LIKE '" + AreaCode.ComboBox.SelectedItem.Value.ToString() + "'", 2, ref pErrorDataTable, "6403", 3, "建设用地总规模（平方公里）") });
            }
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckArea("ZBTZQKB", "ZBDM = '01'", "ZBMJ", treeView1, "ZBFJB", "ZBDM = '01'", "ZBMJ", 2, ref pErrorDataTable, "6402") });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckArea("ZBTZQKB", "ZBDM = '02'", "ZBMJ", treeView1, "ZBFJB", "ZBDM = '02'", "ZBMJ", 1, ref pErrorDataTable, "6402") });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckArea("ZBTZQKB", "ZBDM = '06'", "ZBMJ", treeView1, "ZBFJB", "ZBDM = '06'", "ZBMJ", 0, ref pErrorDataTable, "6402") });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckArea("TDLYJGTZB", "SSJD LIKE '规划年' AND (DLMC LIKE '林地' OR DLMC LIKE '设施农用地' OR DLMC LIKE '河湖水面')", "DLMJ", treeView1, "TDLYJGTZB", "SSJD LIKE '基期年' AND (DLMC LIKE '林地' OR DLMC LIKE '设施农用地' OR DLMC LIKE '河湖水面')", "DLMJ", 1, ref pErrorDataTable, "6403") });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckArea("TDLYJGTZB", "SSJD LIKE '规划年' AND DLMC LIKE '耕地'", "DLMJ", treeView1, "ZBTZQKB", "ZBDM LIKE '01'", "ZBMJ", 1, ref pErrorDataTable, "6403") });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckArea("TDLYJGTZB", "SSJD LIKE '规划年'", "DLMJ", treeView1, "TDLYJGTZB", "SSJD LIKE '基期年'", "DLMJ", 2, ref pErrorDataTable, "6403") });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.CheckArea("GDZBPHB", "BZ LIKE '补划'", "ZBMJ", treeView1, "GDZBPHB", "ZBLX LIKE '建设占用'", "ZBMJ", 1, ref pErrorDataTable, "6403") });
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            this.Invoke(myDelegateAppendTextInfo, new object[] { CDC.JudgeArea(CDC.getArea("TDLYJGTZB","SSJD LIKE '规划年' AND ( DLMC LIKE '耕地')","DLMJ",treeView1)-CDC.getArea("TDLYJGTZB","SSJD LIKE '基期年' AND ( DLMC LIKE '耕地')","DLMJ",treeView1),"TDLYJGTZB",
                CDC.getArea("GDZBPHB","BZ LIKE '补划'","ZBMJ",treeView1)-CDC.getArea("GDZBPHB","BZ LIKE '占用'","ZBMJ",treeView1),"GDZBPHB",2,ref pErrorDataTable,"6403")});
            BindingSource bind = new BindingSource();
            bind.DataSource = pErrorDataTable;
            dgvError.DataSource = bind;
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "非空间数据一致性检查完毕" });
            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\n非空间数据一致性检查完成\r\n完成时间:" + DateTime.Now.ToString() + "\r\n" });
            this.Invoke(myDelegateUpdateBarValue, new object[] { 0 });
        }

        private void buttonCommand4_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            OpenFileDialog opfd1 = new OpenFileDialog();
            opfd1.Filter = "excle文件(*.xls)|*.xlsx|allfile(*.*)|*.*";
            opfd1.Multiselect = false;
            DialogResult diaLres = opfd1.ShowDialog();
            if (diaLres != DialogResult.OK)
                return;
            string path1 = opfd1.FileName;
            //openfiledialog 常规使用
            string pFolder = System.IO.Path.GetDirectoryName(path1);
            string pFileName = System.IO.Path.GetFileName(path1);
            DataTable DT = CDC.GetExcelTable(path1);
            listBox1.Items.Add(path1);
            dgvTable.DataSource = DT;
            tabMapTableView.SelectedTab = tabMapTableView.TabPages[1];
            uiTab2.SelectedTab = uiTab2.TabPages[2];
            //MessageBox.Show(DT.Columns.Count.ToString());
        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem == null) { return; }
            DataTable DT = CDC.GetExcelTable(listBox1.SelectedItem.ToString());
            dgvTable.DataSource = DT;
            tabMapTableView.SelectedTab = tabMapTableView.TabPages[1];
        }

        private void buttonCommand7_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            this.StatisticTable.Rows.Clear();
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "正在禁止建设区规模统计..." });
            progressBar1.Maximum = 2;
            string AppendText = "\r\n禁止建设区规模统计\r\n时间:" + DateTime.Now.ToString();
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            StatisticTable = CDC.getLayerAreaByCity_3(GetLayerByName("JSYDHJBNTGZ2035"), GetLayerByName("STBHKZX"), "GZQLXDM LIKE '04'", "GKDJDM LIKE '01'", "GKDJDM LIKE '02'", GetLayerByName("JCDLYS"));
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            BindingSource bind = new BindingSource();
            bind.DataSource = StatisticTable;
            dgvStastic.DataSource = bind;
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "禁止建设区规模统计完毕" });
            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\n禁止建设区规模统计完成\r\n完成时间:" + DateTime.Now.ToString() + "\r\n" });
            this.uiTab1.SelectedTab = this.uiTab1.TabPages[2];
            this.Invoke(myDelegateUpdateBarValue, new object[] { 0 });
        }

        private void buttonCommand8_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
        {
            this.StatisticTable.Rows.Clear();
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "正在建设用地减量化规模统计..." });
            progressBar1.Maximum = 2;
            string AppendText = "\r\n建设用地减量化规模统计\r\n时间:" + DateTime.Now.ToString();
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            StatisticTable = CDC.getLayerAreaByCity_4(GetLayerByName("JQDLTB"), GetLayerByName("JSYDHJBNTGZ2035"), "DLBM_SX LIKE '2*'", "GZQLXDM LIKE '011' OR GZQLXDM LIKE '012'", GetLayerByName("JCDLYS"));
            this.Invoke(myDelegateUpdateBarValue, new object[] { progressBar1.Value + 1 });
            BindingSource bind = new BindingSource();
            bind.DataSource = StatisticTable;
            dgvStastic.DataSource = bind;
            this.Invoke(myDelegateUpdateUiStatus, new object[] { "建设用地减量化规模统计完毕" });
            this.Invoke(myDelegateAppendTextInfo, new object[] { "\r\n建设用地减量化规模统计完成\r\n完成时间:" + DateTime.Now.ToString() + "\r\n" });
            this.uiTab1.SelectedTab = this.uiTab1.TabPages[2];
            this.Invoke(myDelegateUpdateBarValue, new object[] { 0 });
        }

        private void buttonCommand9_Click(object sender, Janus.Windows.Ribbon.CommandEventArgs e)
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
            sfd.FileName = DateTime.Now.ToString("yyyyMMdd") + "区总规.mdb";
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

                IFeatureLayer mCphFeatureLayer = GetLayerByName("CSKFBJ") as IFeatureLayer;//这是获得要入库的shapefile，获取其FeatureLayer即可
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


                mCphFeatureLayer = GetLayerByName("STBHKZX") as IFeatureLayer;
                IFields pFieldsa = mCphFeatureLayer.FeatureClass.Fields;
                int zdCount = pFieldsa.FieldCount;
                string fileName = "";
                fileName = mCphFeatureLayer.Name;
                ILayer yLayer = GetLayerByName("STBHKZX");
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

                ILayer pLayer1 = GetLayerByName("STBHKZX");
                IFeatureLayer pFeatureLayer1 = pLayer1 as IFeatureLayer;
                IFeatureClass pFeatureClass1 = pFeatureLayer1.FeatureClass;

                IQueryFilter pQueryFilter = new QueryFilterClass();             //SQL Filter 
                pQueryFilter.WhereClause = "GKDJDM LIKE '01' OR GKDJDM LIKE '02'";

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
                        string datasetName = "STBHKZX";//要素数据集的名称
                        IFeatureDataset pCphDataset = CreateFeatureClass(pWorkspace, pCphFeatureClass, datasetName);
                        //3.导入SHP到要素数据集(
                        importToDB(pCphFeatureClass, pWorkspace, pCphDataset, "STBHKZX");
                    }
                }

                mCphFeatureLayer = GetLayerByName("JSYDHJBNTGZ2035") as IFeatureLayer;
                pFieldsa = mCphFeatureLayer.FeatureClass.Fields;
                zdCount = pFieldsa.FieldCount;
                fileName = "";
                fileName = mCphFeatureLayer.Name;
                yLayer = GetLayerByName("JSYDHJBNTGZ2035");
                yFeatureLayer = yLayer as IFeatureLayer;          //获取esriGeometryType,作为参数传入新建shp文件的函数来确定新建类型
                yFeatureClass = yFeatureLayer.FeatureClass;
                fieldname = yFeatureClass.ShapeFieldName;
                yFields = yFeatureClass.Fields;
                ind = yFields.FindField(fieldname);
                yField = yFields.get_Field(ind);

                //IGeometryDef geometryDef = yField.GeometryDef;
                //esriGeometryType type = geometryDef.GeometryType;   //获取esriGeometryType,作为参数传入新建shp文件的
                //IGeometryDefEdit geoDefEdit = geometryDef as IGeometryDefEdit;
                //geoDefEdit.HasZ_2 = false;

                pFeatureLayer = new FeatureLayerClass();          //定义被复制图层和空白shp文件要素和要素类
                pFeatureClass = pFeatureLayer.FeatureClass;
                pFeatureClass = CT.CreateMemoryFeatureClass(mCphFeatureLayer.FeatureClass);

                pLayer1 = GetLayerByName("JSYDHJBNTGZ2035");
                pFeatureLayer1 = pLayer1 as IFeatureLayer;
                pFeatureClass1 = pFeatureLayer1.FeatureClass;

                pQueryFilter = new QueryFilterClass();             //SQL Filter 
                pQueryFilter.WhereClause = "GZQLXDM LIKE '01*' OR GZQLXDM LIKE '02*'";

                pFeatureCursor = pFeatureClass1.Search(pQueryFilter, false);     // 添加要素至要素类并作为一个新涂层展现出来

                pFeature = pFeatureCursor.NextFeature();
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
                        string datasetName = "JSYDHJBNTGZ2035";//要素数据集的名称
                        IFeatureDataset pCphDataset = CreateFeatureClass(pWorkspace, pCphFeatureClass, datasetName);
                        //3.导入SHP到要素数据集(
                        importToDB(pCphFeatureClass, pWorkspace, pCphDataset, "JSYDHJBNTGZ2035");
                    }
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

        private void dgvTable_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {

        }
    }
}