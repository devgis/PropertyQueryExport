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
                //�������Բ�ѯ
                ILayer layer = axMapControl1.get_Layer(0);
                IFeatureLayer featureLayer = layer as IFeatureLayer;
                //��ȡfeatureLayer��featureClass 
                IFeatureClass featureClass = featureLayer.FeatureClass;
                IFeature feature = null;
                IQueryFilter queryFilter = new QueryFilterClass();
                IFeatureCursor featureCusor;
                //���� �Ǹ� ���� ���� ���� ��ϵ
                string where = "1=1";

                #region ��ϲ�ѯ����
                if (!string.IsNullOrEmpty(frmSearchCondation.ResultCondation.TuGang))
                { 
                    where+=string.Format(" and ����='{0}'",frmSearchCondation.ResultCondation.TuGang);
                }
                if (!string.IsNullOrEmpty(frmSearchCondation.ResultCondation.YaGang))
                { 
                    where+=string.Format(" and �Ǹ�='{0}'",frmSearchCondation.ResultCondation.YaGang);
                }
                if (!string.IsNullOrEmpty(frmSearchCondation.ResultCondation.TuLei))
                { 
                    where+=string.Format(" and ����='{0}'",frmSearchCondation.ResultCondation.TuLei);
                }
                if (!string.IsNullOrEmpty(frmSearchCondation.ResultCondation.YaLei))
                { 
                    where+=string.Format(" and ����='{0}'",frmSearchCondation.ResultCondation.YaLei);
                }
                if (!string.IsNullOrEmpty(frmSearchCondation.ResultCondation.TuZu))
                {
                    where+=string.Format(" and ����='{0}'",frmSearchCondation.ResultCondation.TuZu);
                }
                if (!string.IsNullOrEmpty(frmSearchCondation.ResultCondation.TuXi))
                {
                    where += string.Format(" and ��ϵ='{0}'", frmSearchCondation.ResultCondation.TuXi);
                }
                #endregion

                queryFilter.WhereClause = where;
                featureCusor = featureClass.Search(queryFilter, true);
                //search�Ĳ�����һ��Ϊ�����������ڶ���Ϊ�Ƿ��ظ�ִ�С�
                feature = featureCusor.NextFeature();
                if (feature != null)
                {
                    var item = GetTuItem(feature);
                    ShowInfo frmShowInfo = new ShowInfo(item);
                    frmShowInfo.ShowDialog();
                    return;
                }

                MessageBox.Show("û�в�ѯ�����������Ľ����");
            }
        }

        private void menuAbout_Click(object sender, EventArgs e)
        {
            MessageBox.Show("���ް����ĵ���");
        }

        private void menuHelp_Click(object sender, EventArgs e)
        {
            MessageBox.Show("���ǹ��ڣ�");
        }

        private TuItem GetTuItem(IFeature feature)
        {
            TuItem item = new TuItem();

            item.��� = GetFiedValue(feature, "���");

            item.X = GetFiedValue(feature, "X");

            item.Y = GetFiedValue(feature, "Y");

            item.�� = GetFiedValue(feature, "��");

            item.�ص� = GetFiedValue(feature, "�ص�");

            item.���� = GetFiedValue(feature, "���Σ�m��");

            item.�������� = GetFiedValue(feature, "��������");

            item.���������� = GetFiedValue(feature, "����������");

            item.ֲ������
            = GetFiedValue(feature, "ֲ������");

            item.����Ӱ��
            = GetFiedValue(feature, "����Ӱ��");

            item.�����¶�
            = GetFiedValue(feature, "�����¶�");

            item.����ʪ��
            = GetFiedValue(feature, "����ʪ��");

            item.�̽�������
            = GetFiedValue(feature, "�̽�������");

            item.�ǹ̽�����
            = GetFiedValue(feature, "�ǹ̽�����");

            item.��Ч�����
            = GetFiedValue(feature, "��Ч�����");

            item.����ˮ���
            = GetFiedValue(feature, "����ˮ���");

            item.ˮ��
            = GetFiedValue(feature, "ˮ��");

            item.����
            = GetFiedValue(feature, "����");

            item.�Ǹ�
            = GetFiedValue(feature, "�Ǹ�");

            item.����
            = GetFiedValue(feature, "����");

            item.����
            = GetFiedValue(feature, "����");

            item.����
            = GetFiedValue(feature, "����");

            item.��ϵ
            = GetFiedValue(feature, "��ϵ");

            item.�ʵ�
            = GetFiedValue(feature, "�ʵ�");

            item.���������
            = GetFiedValue(feature, "���������");

            item.�ʵ�1
            = GetFiedValue(feature, "�ʵ�1");

            item.����
            = GetFiedValue(feature, "����");

            item.�л���
            = GetFiedValue(feature, "�л���");

            item.�л�̼
            = GetFiedValue(feature, "�л�̼");

            item.ȫ��
            = GetFiedValue(feature, "ȫ��");

            item.��Ч��
            = GetFiedValue(feature, "��Ч��");

            item.ȫ��
            = GetFiedValue(feature, "ȫ��");

            item.��Ч��
            = GetFiedValue(feature, "��Ч��");

            item.ȫ��
            = GetFiedValue(feature, "ȫ��");

            item.��Ч��
            = GetFiedValue(feature, "��Ч��");

            item.CEC
            = GetFiedValue(feature, "CEC");

            item.ȫ��
            = GetFiedValue(feature, "ȫ��");

            item.����̬����
            = GetFiedValue(feature, "����̬����");

            item.ȫ��1
            = GetFiedValue(feature, "ȫ��1");

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
            return string.Empty;//�ֶβ�����
        }
    }
}