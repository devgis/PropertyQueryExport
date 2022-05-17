using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;

using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Controls;
using ESRI.ArcGIS.ADF;
using ESRI.ArcGIS.SystemUI;
using ESRI.ArcGIS.Geodatabase;

namespace SearchApp
{
    public sealed partial class MainForm : Form
    {
        #region class private members
        private IMapControl3 m_mapControl = null;
        private string m_mapDocumentName = string.Empty;
        #endregion

        #region class constructor
        public MainForm()
        {
            InitializeComponent();
        }
        #endregion

        private void MainForm_Load(object sender, EventArgs e)
        {
            //get the MapControl
            m_mapControl = (IMapControl3)axMapControl1.Object;

            //disable the Save menu (since there is no document yet)
            //menuSaveDoc.Enabled = false;
            string file = Path.Combine(Application.StartupPath, "map//map.mxd");
            m_mapControl.LoadMxFile(file);
        }

        #region Main Menu event handlers
        private void menuNewDoc_Click(object sender, EventArgs e)
        {
            //execute New Document command
            ICommand command = new CreateNewDocument();
            command.OnCreate(m_mapControl.Object);
            command.OnClick();
        }

        private void menuOpenDoc_Click(object sender, EventArgs e)
        {
            //execute Open Document command
            ICommand command = new ControlsOpenDocCommandClass();
            command.OnCreate(m_mapControl.Object);
            command.OnClick();
        }

        private void menuSaveDoc_Click(object sender, EventArgs e)
        {
            //execute Save Document command
            if (m_mapControl.CheckMxFile(m_mapDocumentName))
            {
                //create a new instance of a MapDocument
                IMapDocument mapDoc = new MapDocumentClass();
                mapDoc.Open(m_mapDocumentName, string.Empty);

                //Make sure that the MapDocument is not readonly
                if (mapDoc.get_IsReadOnly(m_mapDocumentName))
                {
                    MessageBox.Show("Map document is read only!");
                    mapDoc.Close();
                    return;
                }

                //Replace its contents with the current map
                mapDoc.ReplaceContents((IMxdContents)m_mapControl.Map);

                //save the MapDocument in order to persist it
                mapDoc.Save(mapDoc.UsesRelativePaths, false);

                //close the MapDocument
                mapDoc.Close();
            }
        }

        private void menuSaveAs_Click(object sender, EventArgs e)
        {
            //execute SaveAs Document command
            ICommand command = new ControlsSaveAsDocCommandClass();
            command.OnCreate(m_mapControl.Object);
            command.OnClick();
        }

        private void menuExitApp_Click(object sender, EventArgs e)
        {
            //exit the application
            Application.Exit();
        }
        #endregion

        //listen to MapReplaced evant in order to update the statusbar and the Save menu
        private void axMapControl1_OnMapReplaced(object sender, IMapControlEvents2_OnMapReplacedEvent e)
        {
            //get the current document name from the MapControl
            m_mapDocumentName = m_mapControl.DocumentFilename;

            //if there is no MapDocument, diable the Save menu and clear the statusbar
            if (m_mapDocumentName == string.Empty)
            {
                menuSaveDoc.Enabled = false;
                statusBarXY.Text = string.Empty;
            }
            else
            {
                //enable the Save manu and write the doc name to the statusbar
                menuSaveDoc.Enabled = true;
                statusBarXY.Text = Path.GetFileName(m_mapDocumentName);
            }
        }

        private void axMapControl1_OnMouseMove(object sender, IMapControlEvents2_OnMouseMoveEvent e)
        {
            statusBarXY.Text = string.Format("{0}, {1}  {2}", e.mapX.ToString("#######.##"), e.mapY.ToString("#######.##"), axMapControl1.MapUnits.ToString().Substring(4));
        }

        private void menuPropSearch_Click(object sender, EventArgs e)
        {
            SearchCondation frmSearchCondation = new SearchCondation();
            if (frmSearchCondation.ShowDialog() == DialogResult.OK)
            {
                //进行属性查询
                ILayer layer = axMapControl1.get_Layer(0);
                IFeatureLayer featureLayer = layer as IFeatureLayer;
                //获取featureLayer的featureClass 
                IFeatureClass featureClass = featureLayer.FeatureClass;
                IFeature feature = null;
                IQueryFilter queryFilter = new QueryFilterClass();
                IFeatureCursor featureCusor;
                //土纲 亚纲 土类 亚类 土族 土系
                string where = "1=1";

                #region 组合查询条件
                if (!string.IsNullOrEmpty(frmSearchCondation.ResultCondation.TuGang))
                { 
                    where+=string.Format(" and 土纲='{0}'",frmSearchCondation.ResultCondation.TuGang);
                }
                if (!string.IsNullOrEmpty(frmSearchCondation.ResultCondation.YaGang))
                { 
                    where+=string.Format(" and 亚纲='{0}'",frmSearchCondation.ResultCondation.YaGang);
                }
                if (!string.IsNullOrEmpty(frmSearchCondation.ResultCondation.TuLei))
                { 
                    where+=string.Format(" and 土类='{0}'",frmSearchCondation.ResultCondation.TuLei);
                }
                if (!string.IsNullOrEmpty(frmSearchCondation.ResultCondation.YaLei))
                { 
                    where+=string.Format(" and 亚类='{0}'",frmSearchCondation.ResultCondation.YaLei);
                }
                if (!string.IsNullOrEmpty(frmSearchCondation.ResultCondation.TuZu))
                {
                    where+=string.Format(" and 土族='{0}'",frmSearchCondation.ResultCondation.TuZu);
                }
                if (!string.IsNullOrEmpty(frmSearchCondation.ResultCondation.TuXi))
                {
                    where += string.Format(" and 土系='{0}'", frmSearchCondation.ResultCondation.TuXi);
                }
                #endregion

                queryFilter.WhereClause = where;
                featureCusor = featureClass.Search(queryFilter, true);
                //search的参数第一个为过滤条件，第二个为是否重复执行。
                feature = featureCusor.NextFeature();
                if (feature != null)
                {
                    var item = GetTuItem(feature);
                    ShowInfo frmShowInfo = new ShowInfo(item);
                    frmShowInfo.ShowDialog();
                    return;
                }

                MessageBox.Show("没有查询到符合条件的结果！");
            }
        }

        private void menuAbout_Click(object sender, EventArgs e)
        {
            MessageBox.Show("暂无帮助文档！");
        }

        private void menuHelp_Click(object sender, EventArgs e)
        {
            MessageBox.Show("这是关于！");
        }

        private TuItem GetTuItem(IFeature feature)
        {
            TuItem item = new TuItem();

            item.编号 = GetFiedValue(feature, "编号");

            item.X = GetFiedValue(feature, "X");

            item.Y = GetFiedValue(feature, "Y");

            item.市 = GetFiedValue(feature, "市");

            item.地点 = GetFiedValue(feature, "地点");

            item.海拔 = GetFiedValue(feature, "海拔（m）");

            item.土壤类型 = GetFiedValue(feature, "土壤类型");

            item.土地利用类 = GetFiedValue(feature, "土地利用类");

            item.植被类型
            = GetFiedValue(feature, "植被类型");

            item.人类影响
            = GetFiedValue(feature, "人类影响");

            item.土壤温度
            = GetFiedValue(feature, "土壤温度");

            item.土壤湿度
            = GetFiedValue(feature, "土壤湿度");

            item.固结物质种
            = GetFiedValue(feature, "固结物质种");

            item.非固结物质
            = GetFiedValue(feature, "非固结物质");

            item.有效土层厚
            = GetFiedValue(feature, "有效土层厚");

            item.地下水深度
            = GetFiedValue(feature, "地下水深度");

            item.水质
            = GetFiedValue(feature, "水质");

            item.土纲
            = GetFiedValue(feature, "土纲");

            item.亚纲
            = GetFiedValue(feature, "亚纲");

            item.土类
            = GetFiedValue(feature, "土类");

            item.亚类
            = GetFiedValue(feature, "亚类");

            item.土族
            = GetFiedValue(feature, "土族");

            item.土系
            = GetFiedValue(feature, "土系");

            item.质地
            = GetFiedValue(feature, "质地");

            item.黏土矿物类
            = GetFiedValue(feature, "黏土矿物类");

            item.质地1
            = GetFiedValue(feature, "质地1");

            item.粒径
            = GetFiedValue(feature, "粒径");

            item.有机质
            = GetFiedValue(feature, "有机质");

            item.有机碳
            = GetFiedValue(feature, "有机碳");

            item.全氮
            = GetFiedValue(feature, "全氮");

            item.速效氮
            = GetFiedValue(feature, "速效氮");

            item.全磷
            = GetFiedValue(feature, "全磷");

            item.速效磷
            = GetFiedValue(feature, "速效磷");

            item.全钾
            = GetFiedValue(feature, "全钾");

            item.速效钾
            = GetFiedValue(feature, "速效钾");

            item.CEC
            = GetFiedValue(feature, "CEC");

            item.全铁
            = GetFiedValue(feature, "全铁");

            item.游离态氧化
            = GetFiedValue(feature, "游离态氧化");

            item.全铁1
            = GetFiedValue(feature, "全铁1");

            return item;
        }

        private string GetFiedValue(IFeature feature,string FiledName)
        {
            var index = feature.Fields.FindField(FiledName);
            if (index >= 0)
            {
                try
                {
                    return feature.get_Value(index).ToString();
                }
                catch
                { }
            }
            return string.Empty;//字段不存在
        }
    }
}