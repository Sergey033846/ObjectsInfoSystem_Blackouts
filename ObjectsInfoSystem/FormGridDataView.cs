using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DevExpress.XtraBars.Alerter;

namespace ObjectsInfoSystem
{
    public partial class FormGridDataView : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        // имя таблицы, отображаемой в форме
        public string ds_tblName;        

        public FormGridDataView(string tblName)
        {
            InitializeComponent();

            ds_tblName = tblName;
            gridView1.Columns.Clear();
                        
            switch (tblName)
            {
                case "tblObjects":
                    gridControl1.DataSource = tblObjectsBindingSource;
                    this.Text = "Объекты";
                    break;
                case "tblObjectType":
                    gridControl1.DataSource = tblObjectTypeBindingSource;
                    this.Text = "Типы объектов";
                    break;
                case "tblObjectPropertiesTypes":
                    gridControl1.DataSource = tblObjectPropertiesTypesBindingSource;
                    this.Text = "Свойства типов объектов";
                    break;
                case "tblObjectPropertiesValues":
                    gridControl1.DataSource = tblObjectPropertiesValuesBindingSource;
                    this.Text = "Значения свойств объектов";
                    break;
                case "tblObjectPropCorrValues":
                    gridControl1.DataSource = tblObjectPropCorrValuesBindingSource;
                    this.Text = "Допустимые значения свойств объектов";
                    break;
                //---------------------
                case "tblContractor":
                    gridControl1.DataSource = tblContractorBindingSource;
                    this.Text = "Подрядчики";
                    break;
                case "tblFilial":
                    gridControl1.DataSource = tblFilialBindingSource;
                    this.Text = "Филиалы";
                    break;
                case "tblMunicipality":
                    gridControl1.DataSource = tblMunicipalityBindingSource;
                    this.Text = "Муниципальные образования";
                    break;
                case "tblMapSource":
                    gridControl1.DataSource = tblMapSourceBindingSource;
                    this.Text = "Источники данных";
                    break;
                case "tblPanoramaCoords":
                    gridControl1.DataSource = tblPanoramaCoordsBindingSource;
                    this.Text = "Координаты объектов";
                    break;
                case "tblPanoramaObject":
                    gridControl1.DataSource = tblPanoramaObjectBindingSource;
                    this.Text = "Объекты";
                    break;
                case "tblPanoramaObjSemValues":
                    gridControl1.DataSource = tblPanoramaObjSemValuesBindingSource;
                    this.Text = "Свойства объектов";
                    break;
                case "tblPanoramaSemantic":
                    gridControl1.DataSource = tblPanoramaSemanticBindingSource;
                    this.Text = "Семантика";
                    break;

                // --- ИЭСБК -------
                case "tblIESBKotdelenie":
                    gridControl1.DataSource = tblIESBKotdelenieBindingSource;
                    this.Text = "ИЭСБК отделения";
                    break;
                case "tblIESBKlstype":
                    gridControl1.DataSource = tblIESBKlstypeBindingSource;
                    this.Text = "ИЭСБК тип потребителя";
                    break;
                case "tblIESBKls":
                    gridControl1.DataSource = tblIESBKlsBindingSource;
                    this.Text = "ИЭСБК лицевые счета";
                    break;
                case "tblIESBKlspropvalue":
                    gridControl1.DataSource = tblIESBKlspropvalueBindingSource;
                    this.Text = "ИЭСБК значения свойств";
                    break;
                case "tblIESBKPeriod":
                    gridControl1.DataSource = tblIESBKPeriodBindingSource;
                    this.Text = "Расчетный период";
                    break;
                case "tblIESBKtemplate":
                    gridControl1.DataSource = tblIESBKtemplateBindingSource;
                    this.Text = "ИЭСБК свойства";
                    break;
                case "tblIESBKlsprop":
                    gridControl1.DataSource = tblIESBKlspropBindingSource;
                    this.Text = "ИЭСБК свойства";
                    break;
                // -----------------
            }        
        }

        // пользовательский метод загрузки (обновления) данных
        private void LoadData()
        {
            this.tblObjectsTableAdapter.Fill(this.dataSet1.tblObjects); // необходимо для отображения родительских гридов "+"

            switch (ds_tblName)
            {
                case "tblObjectType":                    
                    this.tblObjectTypeTableAdapter.Fill(this.dataSet1.tblObjectType);
                    break;
                case "tblObjectPropertiesTypes":
                    this.tblObjectPropertiesTypesTableAdapter.Fill(this.dataSet1.tblObjectPropertiesTypes);
                    break;
                case "tblObjectPropertiesValues":
                    this.tblObjectPropertiesValuesTableAdapter.Fill(this.dataSet1.tblObjectPropertiesValues);
                    break;
                case "tblObjectPropCorrValues":
                    this.tblObjectPropCorrValuesTableAdapter.Fill(this.dataSet1.tblObjectPropCorrValues);
                    break;
                //---------------------
                case "tblContractor":
                    this.tblContractorTableAdapter.Fill(this.dataSet1.tblContractor);                    
                    break;
                case "tblFilial":
                    this.tblFilialTableAdapter.Fill(this.dataSet1.tblFilial);
                    break;
                case "tblMunicipality":
                    this.tblMunicipalityTableAdapter.Fill(this.dataSet1.tblMunicipality);
                    break;
                case "tblMapSource":
                    this.tblMapSourceTableAdapter.Fill(this.dataSet1.tblMapSource);
                    break;
                case "tblPanoramaCoords":
                    this.tblPanoramaCoordsTableAdapter.Fill(this.dataSet1.tblPanoramaCoords);
                    break;
                case "tblPanoramaObject":
                    this.tblPanoramaObjectTableAdapter.Fill(this.dataSet1.tblPanoramaObject);
                    break;
                case "tblPanoramaObjSemValues":
                    this.tblPanoramaObjSemValuesTableAdapter.Fill(this.dataSet1.tblPanoramaObjSemValues);
                    break;
                case "tblPanoramaSemantic":
                    this.tblPanoramaSemanticTableAdapter.Fill(this.dataSet1.tblPanoramaSemantic);
                    break;

                // --- ИЭСБК -------
                case "tblIESBKotdelenie":
                    this.tblIESBKotdelenieTableAdapter.Fill(this.dataSetIESBK.tblIESBKotdelenie);
                    break;
                case "tblIESBKlstype":
                    this.tblIESBKlstypeTableAdapter.Fill(this.dataSetIESBK.tblIESBKlstype);
                    break;
                case "tblIESBKls":
                    this.tblIESBKlsTableAdapter.Fill(this.dataSetIESBK.tblIESBKls);                    
                    break;
                case "tblIESBKlspropvalue":
                    this.tblIESBKlspropvalueTableAdapter.Fill(this.dataSetIESBK.tblIESBKlspropvalue);
                    break;
                case "tblIESBKPeriod":
                    this.tblIESBKPeriodTableAdapter.Fill(this.dataSetIESBK.tblIESBKPeriod);
                    break;
                case "tblIESBKtemplate":
                    this.tblIESBKtemplateTableAdapter.Fill(this.dataSetIESBK.tblIESBKtemplate);
                    break;
                case "tblIESBKlsprop":
                    this.tblIESBKlspropTableAdapter.Fill(this.dataSetIESBK.tblIESBKlsprop);
                    break;
                // -----------------
            }
        }

        // пользовательский метод сохранения данных в БД
        private void UpdateData()
        {
            switch (ds_tblName)
            {
                case "tblObjects":
                    this.tblObjectsTableAdapter.Update(this.dataSet1.tblObjects);
                    break;
                case "tblObjectType":
                    this.tblObjectTypeTableAdapter.Update(this.dataSet1.tblObjectType);
                    break;
                case "tblObjectPropertiesTypes":
                    this.tblObjectPropertiesTypesTableAdapter.Update(this.dataSet1.tblObjectPropertiesTypes);
                    break;
                case "tblObjectPropertiesValues":
                    this.tblObjectPropertiesValuesTableAdapter.Update(this.dataSet1.tblObjectPropertiesValues);
                    break;
                case "tblObjectPropCorrValues":
                    this.tblObjectPropCorrValuesTableAdapter.Update(this.dataSet1.tblObjectPropCorrValues);
                    break;
                //---------------------
                case "tblContractor":
                    this.tblContractorTableAdapter.Update(this.dataSet1.tblContractor);
                    break;
                case "tblFilial":
                    this.tblFilialTableAdapter.Update(this.dataSet1.tblFilial);
                    break;
                case "tblMunicipality":
                    this.tblMunicipalityTableAdapter.Update(this.dataSet1.tblMunicipality);
                    break;
                case "tblMapSource":
                    this.tblMapSourceTableAdapter.Update(this.dataSet1.tblMapSource);
                    break;
                case "tblPanoramaCoords":
                    this.tblPanoramaCoordsTableAdapter.Update(this.dataSet1.tblPanoramaCoords);
                    break;
                case "tblPanoramaObject":
                    this.tblPanoramaObjectTableAdapter.Update(this.dataSet1.tblPanoramaObject);
                    break;
                case "tblPanoramaObjSemValues":
                    this.tblPanoramaObjSemValuesTableAdapter.Update(this.dataSet1.tblPanoramaObjSemValues);
                    break;
                case "tblPanoramaSemantic":
                    this.tblPanoramaSemanticTableAdapter.Update(this.dataSet1.tblPanoramaSemantic);
                    break;

                // --- ИЭСБК -------
                case "tblIESBKotdelenie":
                    this.tblIESBKotdelenieTableAdapter.Update(this.dataSetIESBK.tblIESBKotdelenie);
                    break;
                case "tblIESBKlstype":
                    this.tblIESBKlstypeTableAdapter.Update(this.dataSetIESBK.tblIESBKlstype);
                    break;
                case "tblIESBKls":
                    this.tblIESBKlsTableAdapter.Update(this.dataSetIESBK.tblIESBKls);
                    break;
                case "tblIESBKlspropvalue":
                    this.tblIESBKlspropvalueTableAdapter.Update(this.dataSetIESBK.tblIESBKlspropvalue);
                    break;
                case "tblIESBKPeriod":
                    this.tblIESBKPeriodTableAdapter.Update(this.dataSetIESBK.tblIESBKPeriod);
                    break;
                case "tblIESBKtemplate":
                    this.tblIESBKtemplateTableAdapter.Update(this.dataSetIESBK.tblIESBKtemplate);
                    break;
                case "tblIESBKlsprop":
                    this.tblIESBKlspropTableAdapter.Update(this.dataSetIESBK.tblIESBKlsprop);
                    break;
                // -----------------
            }
        }

        private void FormGridDataView_Load(object sender, EventArgs e)
        {            
            LoadData();
        }

        // сохранить
        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            try
            {
                UpdateData();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            splashScreenManager1.CloseWaitForm();

            AlertInfo info = new AlertInfo("Уведомление", "Данные успешно сохранены");
            alertControl1.Show(this, info);
        }

        // экспорт в excel
        private void barButtonItem3_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            this.saveFileDialog1.Title = "Сохранить в Excel";
            this.saveFileDialog1.Filter = "xlsx (*.xlsx)|*.xlsx";
            this.saveFileDialog1.FileName = this.Text;
            if (this.saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    this.gridControl1.ExportToXlsx(this.saveFileDialog1.FileName);                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        // экспорт в csv
        private void barButtonItem4_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            this.saveFileDialog1.Title = "Сохранить в CSV";
            this.saveFileDialog1.Filter = "csv (*.csv)|*.csv";
            this.saveFileDialog1.FileName = this.Text;
            if (this.saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    this.gridControl1.ExportToCsv(this.saveFileDialog1.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        // обновить
        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            LoadData();
            splashScreenManager1.CloseWaitForm();            
        }

        // очистить столбец с присвоением NULL-значений
        private void barButtonItem5_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            FormColumnSelect form1 = new FormColumnSelect();

            form1.Text = "Очистка столбца таблицы (для точек учета)";

            form1.comboBoxEdit1.Properties.Items.Clear();
            form1.comboBoxEdit2.Properties.Items.Clear();
            form1.comboBoxEdit3.Properties.Items.Clear();

            // заполняем поля таблицы
            for (int i = 0; i < this.dataSet1.Tables[ds_tblName].Columns.Count; i++)
                form1.comboBoxEdit1.Properties.Items.Add(this.dataSet1.Tables[ds_tblName].Columns[i].Caption.ToString());
/*
            // выбираем тип проекта
            tblProjectTypeTableAdapter.Fill(dataSet1.tblProjectType);
            for (int i = 0; i < this.dataSet1.tblProjectType.Rows.Count; i++)
                form1.comboBoxEdit2.Properties.Items.Add(this.dataSet1.tblProjectType.Rows[i]["id_caption_prj"].ToString());
                */
            // выбираем филиал
            /*tblFilialTableAdapter.Fill(dataSet1.tblFilial);
            for (int i = 0; i < this.dataSet1.tblProjectType.Rows.Count; i++)
                form1.comboBoxEdit3.Properties.Items.Add(this.dataSet1.tblFilial.Rows[i]["id_caption_filial"].ToString());            
                */
            if (form1.ShowDialog(this) == DialogResult.OK)
            {
                string str0 = null;                

                // пока без филиала и проекта
                //string fexpr = "id_caption_prj = '" + form1.comboBoxEdit2.Text + "'";
                //string fexpr = "id_caption_prj = '" + form1.comboBoxEdit2.Text + "'" + "and id_caption_filial = '" + form1.comboBoxEdit3.Text + "'";
                //DataRow[] mpRows = dataSet1.Tables[ds_tblName].Select(fexpr);
                DataRow[] mpRows = dataSet1.Tables[ds_tblName].Select();

                for (int i = 0; i < mpRows.Length; i++)
                {                    
                    mpRows[i][form1.comboBoxEdit1.Text] = str0;
                }

                // попробовать сделать универсальный
                /*switch (ds_tblName)
                {
                    case "tblMeteringPoint":
                        this.tblMeteringPointTableAdapter.Update(this.dataSet1.tblMeteringPoint);
                        this.tblMeteringPointTableAdapter.Fill(this.dataSet1.tblMeteringPoint);
                        break;
                }*/
                                                                
                this.gridControl1.Refresh();
            }
            else
            {

            }

            form1.Dispose();
        }

        // удалить пробелы в л/с
        private void barButtonItem6_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();

            DataRow[] mpRows = dataSetIESBK.Tables[ds_tblName].Select();
            int maxrow = mpRows.Length;
            for (int i = 0; i < maxrow; i++)
            {
                string code_old = mpRows[i]["codeIESBK"].ToString();
                string code_new = code_old.Replace(" ", "");
                mpRows[i]["codeIESBK"] = code_new;
                tblIESBKlsTableAdapter.Update(mpRows[i]);

                splashScreenManager1.SetWaitFormDescription("Обработка данных (" + (i + 1).ToString() + " из " + maxrow.ToString() + ")");
            }

            splashScreenManager1.CloseWaitForm();
        }
    }
}