using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using ruteo;
using System.IO;
using System.Windows.Forms;
using System.Drawing;
using System.Data;
using System.Globalization;
using System.Runtime.InteropServices;

namespace ruteo
{

    [FormAttribute("UDO_FT_RUTEO")]
    class formRuteo : UDOFormBase
    {
        int progress;
        SAPbobsCOM.Company _SBO;
        public formRuteo()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.btnFormCancel = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.v_fechaEntrega = ((SAPbouiCOM.EditText)(this.GetItem("13_U_E").Specific));
            this.v_fechaEntrega.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.v_fechaEntrega_KeyDownAfter);
            this.v_tipo = ((SAPbouiCOM.ComboBox)(this.GetItem("14_U_Cb").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_8").Specific));
            this.Button1.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button1_ClickBefore);
            this.v_grilla = ((SAPbouiCOM.Grid)(this.GetItem("Item_7").Specific));
            this.v_grilla.DoubleClickAfter += new SAPbouiCOM._IGridEvents_DoubleClickAfterEventHandler(this.v_grilla_DoubleClickAfter);
            this.oMatrix = ((SAPbouiCOM.Matrix)(this.GetItem("0_U_G").Specific));
            this.v_txtChofer = ((SAPbouiCOM.EditText)(this.GetItem("Item_1").Specific));
            this.v_txtChofer.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.v_txtChofer_LostFocusAfter);
            this.v_txtChofer.ClickAfter += new SAPbouiCOM._IEditTextEvents_ClickAfterEventHandler(this.v_txtChofer_ClickAfter);
            this.v_txtChofer.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.v_txtChofer_KeyDownAfter);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_0").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Code = ((SAPbouiCOM.EditText)(this.GetItem("0_U_E").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("Item_3").Specific));
            this.EditText1.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.EditText1_KeyDownAfter);
            this.txtCliente = ((SAPbouiCOM.EditText)(this.GetItem("Item_1").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("Item_4").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.lblCode = ((SAPbouiCOM.StaticText)(this.GetItem("0_U_S").Specific));
            this.lblName = ((SAPbouiCOM.StaticText)(this.GetItem("1_U_S").Specific));
            this.txtName = ((SAPbouiCOM.EditText)(this.GetItem("1_U_E").Specific));
            //    //    inicializar la grilla
            //    SAPbouiCOM.DataTable dt = this.oForm.DataSources.DataTables.Add("dt");
            //    dt.Columns.Add("CHECK", SAPbouiCOM.BoFieldsType.ft_Text);
            //    dt.Columns.Add("Documento", SAPbouiCOM.BoFieldsType.ft_Text);
            //    dt.Columns.Add("Cliente", SAPbouiCOM.BoFieldsType.ft_Text);
            //    dt.Columns.Add("Vencimiento", SAPbouiCOM.BoFieldsType.ft_Date);
            //    dt.Columns.Add("Empleado de venta", SAPbouiCOM.BoFieldsType.ft_Text);
            //    dt.Columns.Add("Total documento", SAPbouiCOM.BoFieldsType.ft_Price);
            //    dt.Columns.Add("Numero interno", SAPbouiCOM.BoFieldsType.ft_Text);
            //    dt.Columns.Add("Parametro", SAPbouiCOM.BoFieldsType.ft_Text);
            //    dt.Columns.Add("Transportista", SAPbouiCOM.BoFieldsType.ft_Text);
            //    dt.Columns.Add("Vehiculo", SAPbouiCOM.BoFieldsType.ft_Text);
            //    dt.Columns.Add("Chofer", SAPbouiCOM.BoFieldsType.ft_Text);
            //    this.v_grilla.DataTable = dt;
            //    this.v_grilla.DataTable.Rows.Add();
            //    this.v_grilla.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
            //    //       this.btnReruteo = ((SAPbouiCOM.Button)(this.GetItem("Item_9").Specific));
            //    //       this.btnReruteo.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.btnReruteo_ClickAfter);
            //    this.oMatrix.Item.Visible = false;
            //    this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_10").Specific));
            //    this.txtmonto = ((SAPbouiCOM.StaticText)(this.GetItem("Item_11").Specific));
            //        agarramos el ultimo codigo pra el ruteo
            //    SAPbobsCOM.Recordset oCode;
            //    oCode = ((SAPbobsCOM.Recordset)(this._SBO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            //    oCode.DoQuery("SELECT MAX(\"DocEntry\")+1 FROM \"@RUTEO\" ");
            //    this.txtmonto.Item.ForeColor = colortxt;
            //    this.Code.Value = oCode.Fields.Item(0).Value.ToString();
            //    this.txtmonto.Caption = "0";
            //    this.txtmonto.Item.FontSize = 15;
            //    int colortxt = System.Drawing.Color.Green.ToArgb();
            //    this.EditText1.Item.Visible = false;
            //    this.EditText2.Item.Visible = false;
            //    this.StaticText1.Item.Visible = false;
            //    this.StaticText2.Item.Visible = false;
            //    this.BtnCancelar = ((SAPbouiCOM.Button)(this.GetItem("Item_12").Specific));
            //    this.BtnCancelar.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.BtnCancelar_ClickAfter);
            //    this.btnFormCancel.Item.Visible = false;
            //    this.txtCliente.Value = "Seleccione chofer...";
            //    this.txtCliente.Item.ForeColor = System.Drawing.Color.DimGray.ToArgb();
            //    this.lblCode.Caption = "Código ruteo";
            //    this.lblName.Item.Visible = false;
            //    this.txtName.Item.Visible = false;
            this.btnCombo = ((SAPbouiCOM.ButtonCombo)(this.GetItem("Item_13").Specific));
            this.btnCombo.PressedAfter += new SAPbouiCOM._IButtonComboEvents_PressedAfterEventHandler(this.btnCombo_PressedAfter);   853494       
            this.OnCustomInitialize();

        }


        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.DataLoadAfter += new SAPbouiCOM.Framework.FormBase.DataLoadAfterHandler(this.Form_DataLoadAfter);
            this.DataLoadBefore += new DataLoadBeforeHandler(this.Form_DataLoadBefore);

        }

        #region DECLARACION DE VARIABLES
        private SAPbouiCOM.EditText v_fechaEntrega;
        private SAPbouiCOM.ComboBox v_tipo;
        private SAPbouiCOM.Matrix oMatrix;
        private SAPbouiCOM.Form oForm;
        SAPbouiCOM.ProgressBar oProgresbar;
        private string v_Consulta = null;
        private SAPbouiCOM.DataTable dt;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.EditText v_txtChofer;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.Grid v_grilla;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.Button Button3;
        private SAPbouiCOM.EditText Code;
        private SAPbouiCOM.Button btnReruteo;
        private SAPbouiCOM.EditText txtCliente;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.StaticText txtmonto;
        private SAPbouiCOM.Button BtnCancelar;
        private SAPbouiCOM.Button btnFormCancel;
        private SAPbouiCOM.StaticText lblCode;
        private SAPbouiCOM.StaticText lblName;
        private SAPbouiCOM.EditText txtName;
        private SAPbouiCOM.ButtonCombo btnCombo;
        private bool flag = false;
        #endregion


        private void EditText0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();


        }

        private void OnCustomInitialize()
        {
            _SBO = conex._sbo;
            //    inicializar la grilla
            SAPbouiCOM.DataTable dt = this.oForm.DataSources.DataTables.Add("dt");
            dt.Columns.Add("CHECK", SAPbouiCOM.BoFieldsType.ft_Text);
            dt.Columns.Add("Documento", SAPbouiCOM.BoFieldsType.ft_Text);
            dt.Columns.Add("Cliente", SAPbouiCOM.BoFieldsType.ft_Text);
            dt.Columns.Add("Vencimiento", SAPbouiCOM.BoFieldsType.ft_Date);
            dt.Columns.Add("Empleado de venta", SAPbouiCOM.BoFieldsType.ft_Text);
            dt.Columns.Add("Total documento", SAPbouiCOM.BoFieldsType.ft_Price);
            dt.Columns.Add("Numero interno", SAPbouiCOM.BoFieldsType.ft_Text);
            dt.Columns.Add("Parametro", SAPbouiCOM.BoFieldsType.ft_Text);
            dt.Columns.Add("Transportista", SAPbouiCOM.BoFieldsType.ft_Text);
            dt.Columns.Add("Vehiculo", SAPbouiCOM.BoFieldsType.ft_Text);
            dt.Columns.Add("Chofer", SAPbouiCOM.BoFieldsType.ft_Text);
            this.v_grilla.DataTable = dt;
            this.v_grilla.DataTable.Rows.Add();
            this.v_grilla.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
            //       this.btnReruteo = ((SAPbouiCOM.Button)(this.GetItem("Item_9").Specific));
            //       this.btnReruteo.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.btnReruteo_ClickAfter);
            this.oMatrix.Item.Visible = false;
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_10").Specific));
            this.txtmonto = ((SAPbouiCOM.StaticText)(this.GetItem("Item_11").Specific));
            //    agarramos el ultimo codigo pra el ruteo
            SAPbobsCOM.Recordset oCode;
            oCode = ((SAPbobsCOM.Recordset)(this._SBO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            oCode.DoQuery("SELECT MAX(\"DocEntry\")+1 FROM \"@RUTEO\" ");
            this.Code.Value = oCode.Fields.Item(0).Value.ToString();
            this.txtmonto.Caption = "0";
            this.txtmonto.Item.FontSize = 15;
            int colortxt = System.Drawing.Color.Green.ToArgb();
            this.txtmonto.Item.ForeColor = colortxt;
            this.EditText1.Item.Visible = false;
            this.EditText2.Item.Visible = false;
            this.StaticText1.Item.Visible = false;
            this.StaticText2.Item.Visible = false;
            this.BtnCancelar = ((SAPbouiCOM.Button)(this.GetItem("Item_12").Specific));
            this.BtnCancelar.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.BtnCancelar_ClickAfter);
            this.btnFormCancel.Item.Visible = false;
            this.txtCliente.Value = "Seleccione chofer...";
            this.txtCliente.Item.ForeColor = System.Drawing.Color.DimGray.ToArgb();
            this.lblCode.Caption = "Código ruteo";
            this.lblName.Item.Visible = false;
            this.txtName.Item.Visible = false;
            this.oMatrix.Item.Visible = true;
            this.Button0.Item.Visible = false;
            this.v_grilla.Item.Visible = false;
            this.btnCombo.ValidValues.Add("Rutear", "Rutear");
            this.btnCombo.ValidValues.Add("Excel", "Excel");
            this.Button3.Item.Enabled = false;
        }
    
        //funcion para cargar la grilla
        private void v_fechaEntrega_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            this.oMatrix.Item.Visible = false;
            this.v_grilla.Item.Visible = true;
            //consultamos si se presiono TAB
            if (pVal.CharPressed == (char)9)
            {
                string v_tipoText = v_tipo.Selected.Description.ToString();
                string v_fecha = v_fechaEntrega.Value;
                //DateTime fecha = DateTime.Parse(v_fecha);
                //string fecha_v = fecha.ToString("yyyyMMdd");
                //en caso de que no se haya seleccionado nada que tire un mensaje           
                if (v_tipoText.Equals("Seleccionar"))
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Seleccione una sucursal", 1, "Ok");
                    return;
                }
                //consultar los querys en caso de que se haya seleccionado un tipo

                if (v_tipoText.Equals("ITAUGUA"))
                {
                    v_Consulta = "SELECT T0.\"DocNum\",T0.\"CardName\", T0.\"DocDueDate\", T0.\"U_Vehiculo\",T0.\"U_Chofer\",T1.\"SlpName\", T0.\"DocTotal\", T0.\"DocEntry\" , " +
                                 "case when MIN(T3.\"WhsCode\") = 'ITG-HAM' THEN 'Hamburguesas'  when  MIN(T3.\"WhsCode\") = 'ITG-EMB' THEN 'Embutidos' when MIN(T3.\"WhsCode\") = 'ITG-PYP' THEN 'Papas y Pizas' ELSE 'Alimentos Secos' END AS PARAMETRO "+
                                 "FROM ORDR T0 "+
                                 "JOIN \"OSLP\" T1 on  T0.\"SlpCode\"=T1.\"SlpCode\" "+
                                 "inner join \"NNM1\" T2  ON T0.\"Series\" = T2.\"Series\" "+
                                 "INNER JOIN \"RDR1\" T3 ON T0.\"DocEntry\" = T3.\"DocEntry\" "+
                                 "WHERE T0.\"DocDueDate\" = '"+ v_fecha + "' and  T2.\"SeriesName\" LIKE '017%' AND T3.\"WhsCode\" in ('ITG-HAM', 'ITG-EMB', 'ITG-SEC', 'ITG-PYP') AND  T0.\"CANCELED\" = 'N' "+
                                 "GROUP BY T0.\"DocNum\",T0.\"CardName\", T0.\"DocDueDate\", T0.\"U_Vehiculo\",T0.\"U_Chofer\",T1.\"SlpName\",T0.\"DocTotal\", T0.\"DocEntry\" ";
                }
                if (v_tipoText.Equals("CDE"))
                {
                    v_Consulta = "SELECT T0.\"DocNum\",T0.\"CardName\", T0.\"DocDueDate\", T0.\"U_Vehiculo\",T0.\"U_Chofer\",T1.\"SlpName\", T0.\"DocTotal\", T0.\"DocEntry\" " +                                 
                                 "FROM ORDR T0 " +
                                 "JOIN \"OSLP\" T1 on  T0.\"SlpCode\"=T1.\"SlpCode\" " +                                 
                                 "WHERE T0.\"DocDueDate\" = '" + v_fecha + "' and  T0.\"Series\" = '123'  AND  T0.\"CANCELED\" = 'N' " +
                                 "ORDER BY T0.\"CardName\", T0.\"U_Chofer\" ";
                }
                if (v_tipoText.Equals("CDI"))
                {
                    v_Consulta = "SELECT T0.\"DocNum\",T0.\"CardName\", T0.\"DocDueDate\", T0.\"U_Vehiculo\",T0.\"U_Chofer\",T1.\"SlpName\", T0.\"DocTotal\", T0.\"DocEntry\",'' " +
                                 "FROM ORDR T0 " +
                                 "JOIN \"OSLP\" T1 on  T0.\"SlpCode\"=T1.\"SlpCode\" " +
                                 "WHERE T0.\"DocDueDate\" = '" + v_fecha + "' and  T0.\"Series\" = '525'  AND  T0.\"CANCELED\" = 'Y' AND T0.\"CardCode\"='CI0000559' " +
                                 "ORDER BY T0.\"CardName\", T0.\"U_Chofer\" ";
                }
                //consultamos a la base de datos
                SAPbobsCOM.Recordset oConsulta;
                oConsulta = (SAPbobsCOM.Recordset)_SBO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oConsulta.DoQuery(v_Consulta);
                int v_can = oConsulta.RecordCount;
                //instanciamos la matriz
                SAPbouiCOM.DBDataSource source = oForm.DataSources.DBDataSources.Item("@RUTEODET");
                oMatrix.FlushToDataSource();
                source.Clear();
                int v_filaMatrix = 0;
                int v_canInicio = 1;
                oProgresbar = SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.CreateProgressBar("Cargando", v_can, true);
                
                while (!oConsulta.EoF)
                {
                    oProgresbar.Value = oProgresbar.Value + 1;
                    oProgresbar.Text = "Cargando Datos " + v_canInicio + "/" + v_can.ToString();
                    string documento = oConsulta.Fields.Item(0).Value.ToString();
                    string cliente = oConsulta.Fields.Item(1).Value.ToString();
                    string vencimiento = oConsulta.Fields.Item(2).Value.ToString();
                    DateTime venci = DateTime.Parse(vencimiento);
                    string vencimiento_V = venci.ToString("yyyyMMdd");
                    string empleado_venta = oConsulta.Fields.Item(5).Value.ToString();
                    string total = oConsulta.Fields.Item(6).Value.ToString();
                    string interno = oConsulta.Fields.Item(7).Value.ToString();
                    string parametro = oConsulta.Fields.Item(8).Value.ToString();

                    source.InsertRecord(source.Size);
                    source.Offset = source.Size - 1;
                    source.SetValue("U_Documento", v_filaMatrix, documento);
                    source.SetValue("U_Cliente", v_filaMatrix, cliente);
                    source.SetValue("U_Vencimiento", v_filaMatrix, vencimiento_V);
                    source.SetValue("U_Emp_venta", v_filaMatrix, empleado_venta);
                    source.SetValue("U_Total", v_filaMatrix, total);
                    source.SetValue("U_Num_interno", v_filaMatrix, interno);
                    source.SetValue("U_Parametro", v_filaMatrix, parametro);
                    oMatrix.LoadFromDataSource();
                    

                    //cargar filas de la grilla
                    v_grilla.DataTable.SetValue("Documento", v_filaMatrix, documento);
                    v_grilla.DataTable.SetValue("Cliente", v_filaMatrix, cliente);
                    v_grilla.DataTable.SetValue("Vencimiento", v_filaMatrix, vencimiento_V);
                    v_grilla.DataTable.SetValue("Empleado de venta", v_filaMatrix, empleado_venta);
                    v_grilla.DataTable.SetValue("Total documento", v_filaMatrix, total);
                    v_grilla.DataTable.SetValue("Numero interno", v_filaMatrix, interno);
                    v_grilla.DataTable.SetValue("Parametro", v_filaMatrix, parametro);
                    //v_grilla.DataTable.SetValue("Transportista", v_filaMatrix, documento);
                    //v_grilla.DataTable.SetValue("Vehiculo", v_filaMatrix, documento);
                    //v_grilla.DataTable.SetValue("Chofer", v_filaMatrix, documento);
                    v_grilla.DataTable.Rows.Add();

                    v_filaMatrix++;
                    oConsulta.MoveNext();
                    progress = progress + 1;
                    v_canInicio++;
                }
                oProgresbar.Stop();
            }
            

        }
      
        //funcion para descargar a excel
        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;            

            Microsoft.Office.Interop.Excel.Application aplicacion;
            Microsoft.Office.Interop.Excel.Workbook libro;
            Microsoft.Office.Interop.Excel.Worksheet hoja;
            Microsoft.Office.Interop.Excel.Worksheet hoja2;
            Microsoft.Office.Interop.Excel.Range rango;
            object misvalue = System.Reflection.Missing.Value;
            SAPbouiCOM.Matrix grilla = (SAPbouiCOM.Matrix)oForm.Items.Item("0_U_G").Specific;

            try
            {
                //configuramos los elementos para el excel
                aplicacion = new Microsoft.Office.Interop.Excel.Application();
                aplicacion.Visible = false;
                libro = (Microsoft.Office.Interop.Excel.Workbook)(aplicacion.Workbooks.Add(""));
                hoja = (Microsoft.Office.Interop.Excel.Worksheet)libro.ActiveSheet;
                //agregamos los titulos al excel
                hoja.Cells[1, 1] = "DOCUMENTO";
                hoja.Cells[1, 2] = "CLIENTE";
                hoja.Cells[1, 3] = "VENCIMIENTO";
                hoja.Cells[1, 4] = "EMPLEADO DE VENTA";
                hoja.Cells[1, 5] = "TOTAL DOCUMENTO";
                hoja.Cells[1, 6] = "NUMERO INTERNO";
                hoja.Cells[1, 7] = "PARAMETRO";
                hoja.Cells[1, 8] = "TRANSPORTISTA";
                hoja.Cells[1, 9] = "VEHICULO";
                hoja.Cells[1, 10] = "CHOFER";
                //ponemos en negrita los titulos
                hoja.Range["A1", "J1"].Font.Bold = true;

                int fila = 1;
                int filacelda = 2;
                int filamatrix = 1;
                int countgrid = grilla.RowCount;
                //recorremos la grilla
                while (fila <= countgrid)
                {
                    SAPbouiCOM.EditText oItem1 = (SAPbouiCOM.EditText)grilla.Columns.Item("C_0_1").Cells.Item(filamatrix).Specific;
                    SAPbouiCOM.EditText oItem2 = (SAPbouiCOM.EditText)grilla.Columns.Item("C_0_2").Cells.Item(filamatrix).Specific;
                    SAPbouiCOM.EditText oItem3 = (SAPbouiCOM.EditText)grilla.Columns.Item("C_0_3").Cells.Item(filamatrix).Specific;
                    SAPbouiCOM.EditText oItem4 = (SAPbouiCOM.EditText)grilla.Columns.Item("C_0_4").Cells.Item(filamatrix).Specific;
                    SAPbouiCOM.EditText oItem5 = (SAPbouiCOM.EditText)grilla.Columns.Item("C_0_5").Cells.Item(filamatrix).Specific;
                    SAPbouiCOM.EditText oItem6 = (SAPbouiCOM.EditText)grilla.Columns.Item("C_0_6").Cells.Item(filamatrix).Specific;
                    SAPbouiCOM.EditText oItem7 = (SAPbouiCOM.EditText)grilla.Columns.Item("C_0_7").Cells.Item(filamatrix).Specific;
                    SAPbouiCOM.EditText oItem8 = (SAPbouiCOM.EditText)grilla.Columns.Item("C_0_8").Cells.Item(filamatrix).Specific;
                    SAPbouiCOM.EditText oItem9 = (SAPbouiCOM.EditText)grilla.Columns.Item("C_0_9").Cells.Item(filamatrix).Specific;
                    SAPbouiCOM.EditText oItem10 = (SAPbouiCOM.EditText)grilla.Columns.Item("C_0_10").Cells.Item(filamatrix).Specific;

                    string v_item1 = oItem1.Value;
                    string v_item2 = oItem2.Value;
                    string v_item3 = oItem3.Value;
                    string v_item4 = oItem4.Value;
                    string v_item5 = oItem5.Value;
                    string v_item6 = oItem6.Value;
                    string v_item7 = oItem7.Value;
                    string v_item8 = oItem8.Value;
                    string v_item9 = oItem9.Value;
                    string v_item10 = oItem10.Value;

                    hoja.Cells[filacelda, 1] = v_item1;// grilla.DataTable.GetValue(0, filamatrix);
                    hoja.Cells[filacelda, 2] = v_item2;// grilla.DataTable.GetValue(1, filamatrix);
                    hoja.Cells[filacelda, 3] = v_item3;// grilla.DataTable.GetValue(2, filamatrix);
                    hoja.Cells[filacelda, 4] = v_item4;// grilla.DataTable.GetValue(3, filamatrix);
                    hoja.Cells[filacelda, 5] = v_item5;// grilla.DataTable.GetValue(4, filamatrix);
                    hoja.Cells[filacelda, 6] = v_item6;// grilla.DataTable.GetValue(5, filamatrix);
                    hoja.Cells[filacelda, 7] = v_item7;// grilla.DataTable.GetValue(6, filamatrix);
                    hoja.Cells[filacelda, 8] = v_item8;// grilla.DataTable.GetValue(7, filamatrix);
                    hoja.Cells[filacelda, 9] = v_item9;// grilla.DataTable.GetValue(8, filamatrix);
                    hoja.Cells[filacelda, 10] = v_item10;// grilla.DataTable.GetValue(9, filamatrix);

                    filacelda = filacelda + 1;
                    filamatrix = filamatrix + 1;
                    fila = fila + 1;
                }
                //creamos una carpeta en el escritorio para guardar el excel
                string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string CarpEscr = path + "\\RUTEO";
                if (!Directory.Exists(CarpEscr))
                {
                    Directory.CreateDirectory(CarpEscr);
                }
                string v_texto = v_tipo.Selected.Description.ToString();
                aplicacion.Visible = false;
                aplicacion.UserControl = false;
                string archivo = CarpEscr + "\\RUTEO_"+ v_texto+"_" + DateTime.Now.Hour.ToString("D2") + "" + DateTime.Now.Minute.ToString("D2") + "" + DateTime.Now.Second.ToString("D2") + ".xls";
                libro.SaveAs(archivo);
                libro.Close();
                aplicacion.Quit();
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Exportado con éxito", 1, "Ok");

            }
            catch (Exception e)
            {

            }

        }

        private void EditText1_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            if (pVal.CharPressed == (char)9)
            {
                try
                {
                    //agarramos el codigo primario
                    string v_code = Code.Value;
                    //consultamos a la base
                    SAPbobsCOM.Recordset oConsulta2;
                    oConsulta2 = (SAPbobsCOM.Recordset)_SBO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oConsulta2.DoQuery("SELECT \"U_Documento\",\"U_Cliente\",\"U_Vencimiento\",\"U_Emp_venta\",\"U_Total\",\"U_Num_interno\",\"U_Parametro\",\"U_Transportista\",\"U_Vehiculo\",\"U_Chofer\" FROM \"@RUTEODET\" WHERE \"Code\"='" + v_code + "' ");
                    //cargamos la grilla
                    int v_fila = 0;
                    int v_filacolor = 1;

                    while (!oConsulta2.EoF)
                    {
                        string documento = oConsulta2.Fields.Item(0).Value.ToString();
                        string cliente = oConsulta2.Fields.Item(1).Value.ToString();
                        string vencimiento = oConsulta2.Fields.Item(2).Value.ToString();
                        DateTime venci = DateTime.Parse(vencimiento);
                        string vencimiento_V = venci.ToString("yyyyMMdd");
                        string empleado_venta = oConsulta2.Fields.Item(3).Value.ToString();
                        string total = oConsulta2.Fields.Item(4).Value.ToString();
                        string interno = oConsulta2.Fields.Item(5).Value.ToString();
                        string parametro = oConsulta2.Fields.Item(6).Value.ToString();
                        string trans = oConsulta2.Fields.Item(7).Value.ToString();
                        string chapa = oConsulta2.Fields.Item(8).Value.ToString();
                        string chofer = oConsulta2.Fields.Item(9).Value.ToString();

                        //cargar filas de la grilla
                        v_grilla.DataTable.SetValue("Documento", v_fila, documento);
                        v_grilla.DataTable.SetValue("Cliente", v_fila, cliente);
                        v_grilla.DataTable.SetValue("Vencimiento", v_fila, vencimiento_V);
                        v_grilla.DataTable.SetValue("Empleado de venta", v_fila, empleado_venta);
                        v_grilla.DataTable.SetValue("Total documento", v_fila, total);
                        v_grilla.DataTable.SetValue("Numero interno", v_fila, interno);
                        v_grilla.DataTable.SetValue("Parametro", v_fila, parametro);
                        v_grilla.DataTable.SetValue("Transportista", v_fila, trans);
                        v_grilla.DataTable.SetValue("Vehiculo", v_fila, chapa);
                        v_grilla.DataTable.SetValue("Chofer", v_fila, chofer);
                        if (!string.IsNullOrEmpty(trans))
                        {
                            int color = Color.LightGreen.ToArgb();
                            v_grilla.CommonSetting.SetRowBackColor(v_filacolor, color);
                        }
                        v_grilla.DataTable.Rows.Add();

                        v_fila++;
                        v_filacolor++;
                        oConsulta2.MoveNext();
                    }
                }
                catch (Exception e)
                {

                }

            }

        }

        //funcion para cargar los choferes a la grilla
        private void Button1_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            this.Button3.Item.Enabled = false;
            string v_chofer = v_txtChofer.Value.ToString();
            string v_transp = null;
            string v_chapa = null;
            if (string.IsNullOrEmpty(v_chofer))
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Seleccione un chofer para el ruteo", 1, "Ok");
                return;
            }

            //buscamos los el transportista y vehiculo
            SAPbobsCOM.Recordset oRecord;
            oRecord = (SAPbobsCOM.Recordset)_SBO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecord.DoQuery("SELECT \"U_TransCod\",\"U_Chapa\" FROM \"@CHOF_RUTEO\" WHERE \"Code\"='" + v_chofer + "' ");
            while (!oRecord.EoF)
            {
                v_transp = oRecord.Fields.Item(0).Value.ToString();
                v_chapa = oRecord.Fields.Item(1).Value.ToString();

                oRecord.MoveNext();
            }

            //throw new System.NotImplementedException();
            int v_rowsGrid = v_grilla.Rows.Count;
            int v_fila = 0;
            int v_filaColor = 1;
            //SAPbobsCOM.Documents oOrden;
            //oOrden = (SAPbobsCOM.Documents)_SBO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

            while (v_fila < v_rowsGrid)
            {
                string v_check = v_grilla.DataTable.GetValue("CHECK", v_fila).ToString();
                if (v_check.Equals("Y"))
                {
                    //cargar chofer a los pedidos - grilla
                    v_grilla.DataTable.SetValue("Transportista", v_fila, v_transp);
                    v_grilla.DataTable.SetValue("Vehiculo", v_fila, v_chapa);
                    v_grilla.DataTable.SetValue("Chofer", v_fila, v_chofer);
                    v_grilla.DataTable.SetValue("CHECK", v_fila, "");
                    int colorGrilla = Color.LightGray.ToArgb();
                    v_grilla.CommonSetting.SetRowBackColor(v_filaColor, colorGrilla);
                    //grabar datos en el pedido
                    string pedido = v_grilla.DataTable.GetValue("Numero interno", v_fila).ToString();
                    string v_parametro = v_grilla.DataTable.GetValue("Parametro", v_fila).ToString();
                    //try
                    //{
                    //    int v_pedido = int.Parse(pedido);
                    //    //string v_numPedido = oBusqueda.Fields.Item(4).Value.ToString();
                    //    if (oOrden.GetByKey(v_pedido))
                    //    {
                    //        oOrden.UserFields.Fields.Item("U_Trans").Value = v_transp;
                    //        oOrden.UserFields.Fields.Item("U_Vehiculo").Value = v_chapa;
                    //        oOrden.UserFields.Fields.Item("U_Chofer").Value = v_chofer;
                    //        oOrden.UserFields.Fields.Item("U_par").Value = v_parametro;
                    //        int up = oOrden.Update();
                    //        if (up != 0)
                    //        {

                    //            int color2 = Color.Blue.ToArgb();
                    //            v_grilla.CommonSetting.SetRowBackColor(v_filaColor, color2);
                    //        }
                    //        else
                    //        {

                    //            int color = Color.LightGreen.ToArgb();
                    //            v_grilla.CommonSetting.SetRowBackColor(v_filaColor, color);
                    //        }

                    //    }
                    //}
                    //catch
                    //{
                    //    int color2 = Color.Blue.ToArgb();
                    //    v_grilla.CommonSetting.SetRowBackColor(v_filaColor, color2);
                    //}
                    //cargar chofer a los pedidos - matrix
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item(8).Cells.Item(v_filaColor).Specific).Value = v_transp;
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item(9).Cells.Item(v_filaColor).Specific).Value = v_chapa;
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item(10).Cells.Item(v_filaColor).Specific).Value = v_chofer;

                }
                v_fila++;
                v_filaColor++;
            }
            v_txtChofer.Value = "";
            //int res = SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("¿Desea seguir ruteando?", 2, "SI", "NO");
            //if (res == 2)
            //{
            //    v_grilla.Item.Visible = false;
            //    oMatrix.Item.Visible = true;               
            //}



        }

        //funcion para actualizar los pedidos
        private void rutearPedidos(string pedido, string trans, string chapa, string chofer)
        {
            //string v_newKey =  _SBO.GetNewObjectKey();
            string v_ListError = "No se actualizo el siguiente pedido: ";
            
            SAPbobsCOM.Documents oOrden;
            oOrden = (SAPbobsCOM.Documents)_SBO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

            int v_pedido = int.Parse(pedido);
            //try
            //{
            //    int v_pedido = int.Parse(pedido);
            //    //string v_numPedido = oBusqueda.Fields.Item(4).Value.ToString();
            //    if (oOrden.GetByKey(v_pedido))
            //    {
            //        oOrden.UserFields.Fields.Item("U_Trans").Value = v_transp;
            //        oOrden.UserFields.Fields.Item("U_Vehiculo").Value = v_chapa;
            //        oOrden.UserFields.Fields.Item("U_Chofer").Value = v_chofer;
            //        oOrden.UserFields.Fields.Item("U_par").Value = v_parametro;
            //        int up = oOrden.Update();
            //        if (up != 0)
            //        {

            //            int color2 = Color.Blue.ToArgb();
            //            v_grilla.CommonSetting.SetRowBackColor(v_filaColor, color2);
            //        }
            //        else
            //        {

            //            int color = Color.LightGreen.ToArgb();
            //            v_grilla.CommonSetting.SetRowBackColor(v_filaColor, color);
            //        }

            //    }
            //}
            //catch
            //{
            //    int color2 = Color.Blue.ToArgb();
            //    v_grilla.CommonSetting.SetRowBackColor(v_filaColor, color2);
            //}

        }

        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {           

        }

        //total en valores de los pedidos
        private void v_txtChofer_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            if (pVal.CharPressed == (char)9)
            {
                NumberFormatInfo nfi = new CultureInfo("en-US",false).NumberFormat;
                string v_clinte = txtCliente.Value;
                decimal v_monto = 0;
                int canRow = v_grilla.Rows.Count;
                int fila = 1;
                int filagrilla = 0;
                while (fila < canRow)
                {
                    string G_cliente = v_grilla.DataTable.GetValue("Chofer", filagrilla).ToString();
                    if (G_cliente.Equals(v_clinte))
                    {
                        string G_monto = v_grilla.DataTable.GetValue("Total documento", filagrilla).ToString();
                        v_monto = v_monto + decimal.Parse(G_monto);
                        txtmonto.Caption = v_monto.ToString("N",nfi);
                    }
                    fila++;
                    filagrilla++;
                }
                if (v_monto == 0)
                {
                    txtmonto.Caption = v_monto.ToString();
                }
            }
        }

        //funcion del boton cancelar
        private void BtnCancelar_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            string v_trans = null;
            if (pVal.FormMode == 3)
            {
                int cantGrilla = v_grilla.Rows.Count;
                int filas = 0;
                if (cantGrilla > 1)
                {
                    while (filas < cantGrilla)
                    {
                        v_trans = v_grilla.DataTable.GetValue("Transportista", filas).ToString();
                        if (!string.IsNullOrEmpty(v_trans))
                        {
                            SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Debe guardar el ruteo antes de salir del formulario!!", 1, "Ok");
                            return;
                        }
                        filas++;
                    }
                    if (string.IsNullOrEmpty(v_trans))
                    {
                        this.oForm.Close();
                    }
                }
                else
                {
                    this.oForm.Close();
                }
            }
            else
            {
                this.oForm.Close();
            }
            
            
        }

        //funcion para borrar el placeholder de chofer
        private void v_txtChofer_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            string v_texto = txtCliente.Value;
            if(v_texto.Equals("Seleccione chofer..."))
            {
                txtCliente.Value = "";
            }

        }

        //funcion para poner el placeholde rde chofer
        private void v_txtChofer_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            string v_texto = txtCliente.Value;
            if (v_texto.Equals(""))
            {
                txtCliente.Value = "Seleccione chofer...";
            }
        }

        private void Form_DataLoadBefore(ref SAPbouiCOM.BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //throw new System.NotImplementedException();

        }

        //funcion para abrir el form de pedidos (orden de venta)
        private void v_grilla_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            string pp =  pVal.ColUID;            
            if (pVal.ColUID.Equals("Documento"))
            {
                string v_documento = v_grilla.DataTable.GetValue("Numero interno", pVal.Row).ToString();
                SAPbouiCOM.Framework.Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_Order,"",v_documento);
            }
        }     

        //funcion de combo d eboton
        private void btnCombo_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (btnCombo.Caption.Equals("Rutear"))
            {
                this.Button3.Item.Enabled = true;
                //instanciamos el objeto
                SAPbobsCOM.Documents oOrden;
                oOrden = (SAPbobsCOM.Documents)_SBO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                //contamos la cantidad de lineas a recorrer
                int cantGrilla = v_grilla.Rows.Count-1;
                int fila = 0;
                int filacolor = 1;
                while (fila < cantGrilla)
                {
                    string pedido = v_grilla.DataTable.GetValue("Numero interno", fila).ToString();
                    string v_transp = v_grilla.DataTable.GetValue("Transportista", fila).ToString();
                    string v_chapa = v_grilla.DataTable.GetValue("Vehiculo", fila).ToString();
                    string v_chofer = v_grilla.DataTable.GetValue("Chofer", fila).ToString();
                    string v_parametro = v_grilla.DataTable.GetValue("Parametro", fila).ToString();
                    try
                    {
                        int v_pedido = int.Parse(pedido);
                        //string v_numPedido = oBusqueda.Fields.Item(4).Value.ToString();
                        if (oOrden.GetByKey(v_pedido))
                        {
                            oOrden.UserFields.Fields.Item("U_Trans").Value = v_transp;
                            oOrden.UserFields.Fields.Item("U_Vehiculo").Value = v_chapa;
                            oOrden.UserFields.Fields.Item("U_Chofer").Value = v_chofer;
                            oOrden.UserFields.Fields.Item("U_par").Value = v_parametro;
                            if (!v_parametro.Equals("Alimentos Secos"))
                            {
                                oOrden.UserFields.Fields.Item("U_LEYENDA").Value = "2";
                            }                            
                            oOrden.NumAtCard = "ITAUGUA";
                            int up = oOrden.Update();
                            if (up != 0)
                            {

                                int color2 = Color.LightBlue.ToArgb();
                                v_grilla.CommonSetting.SetRowBackColor(filacolor, color2);
                            }
                            else
                            {

                                int color = Color.LightGreen.ToArgb();
                                v_grilla.CommonSetting.SetRowBackColor(filacolor, color);
                            }

                        }
                    }
                    catch
                    {
                        int color2 = Color.LightBlue.ToArgb();
                        v_grilla.CommonSetting.SetRowBackColor(filacolor, color2);
                    }
                    fila++;
                    filacolor++;
                }
            }
            if (btnCombo.Caption.Equals("Excel"))
            {
                Microsoft.Office.Interop.Excel.Application aplicacion;
                Microsoft.Office.Interop.Excel.Workbook libro;
                Microsoft.Office.Interop.Excel.Worksheet hoja;
                Microsoft.Office.Interop.Excel.Worksheet hoja2;
                Microsoft.Office.Interop.Excel.Range rango;
                object misvalue = System.Reflection.Missing.Value;
                SAPbouiCOM.Matrix grilla = (SAPbouiCOM.Matrix)oForm.Items.Item("0_U_G").Specific;

                try
                {
                    //configuramos los elementos para el excel
                    aplicacion = new Microsoft.Office.Interop.Excel.Application();
                    aplicacion.Visible = false;
                    libro = (Microsoft.Office.Interop.Excel.Workbook)(aplicacion.Workbooks.Add(""));
                    hoja = (Microsoft.Office.Interop.Excel.Worksheet)libro.ActiveSheet;
                    //agregamos los titulos al excel
                    hoja.Cells[1, 1] = "DOCUMENTO";
                    hoja.Cells[1, 2] = "CLIENTE";
                    hoja.Cells[1, 3] = "VENCIMIENTO";
                    hoja.Cells[1, 4] = "EMPLEADO DE VENTA";
                    hoja.Cells[1, 5] = "TOTAL DOCUMENTO";
                    hoja.Cells[1, 6] = "NUMERO INTERNO";
                    hoja.Cells[1, 7] = "PARAMETRO";
                    hoja.Cells[1, 8] = "TRANSPORTISTA";
                    hoja.Cells[1, 9] = "VEHICULO";
                    hoja.Cells[1, 10] = "CHOFER";
                    //ponemos en negrita los titulos
                    hoja.Range["A1", "J1"].Font.Bold = true;

                    int fila = 1;
                    int filacelda = 2;
                    int filamatrix = 1;
                    int countgrid = grilla.RowCount;
                    //recorremos la grilla
                    while (fila <= countgrid)
                    {
                        SAPbouiCOM.EditText oItem1 = (SAPbouiCOM.EditText)grilla.Columns.Item("C_0_1").Cells.Item(filamatrix).Specific;
                        SAPbouiCOM.EditText oItem2 = (SAPbouiCOM.EditText)grilla.Columns.Item("C_0_2").Cells.Item(filamatrix).Specific;
                        SAPbouiCOM.EditText oItem3 = (SAPbouiCOM.EditText)grilla.Columns.Item("C_0_3").Cells.Item(filamatrix).Specific;
                        SAPbouiCOM.EditText oItem4 = (SAPbouiCOM.EditText)grilla.Columns.Item("C_0_4").Cells.Item(filamatrix).Specific;
                        SAPbouiCOM.EditText oItem5 = (SAPbouiCOM.EditText)grilla.Columns.Item("C_0_5").Cells.Item(filamatrix).Specific;
                        SAPbouiCOM.EditText oItem6 = (SAPbouiCOM.EditText)grilla.Columns.Item("C_0_6").Cells.Item(filamatrix).Specific;
                        SAPbouiCOM.EditText oItem7 = (SAPbouiCOM.EditText)grilla.Columns.Item("C_0_7").Cells.Item(filamatrix).Specific;
                        SAPbouiCOM.EditText oItem8 = (SAPbouiCOM.EditText)grilla.Columns.Item("C_0_8").Cells.Item(filamatrix).Specific;
                        SAPbouiCOM.EditText oItem9 = (SAPbouiCOM.EditText)grilla.Columns.Item("C_0_9").Cells.Item(filamatrix).Specific;
                        SAPbouiCOM.EditText oItem10 = (SAPbouiCOM.EditText)grilla.Columns.Item("C_0_10").Cells.Item(filamatrix).Specific;

                        string v_item1 = oItem1.Value;
                        string v_item2 = oItem2.Value;
                        string v_item3 = oItem3.Value;
                        string v_item4 = oItem4.Value;
                        string v_item5 = oItem5.Value;
                        string v_item6 = oItem6.Value;
                        string v_item7 = oItem7.Value;
                        string v_item8 = oItem8.Value;
                        string v_item9 = oItem9.Value;
                        string v_item10 = oItem10.Value;

                        hoja.Cells[filacelda, 1] = v_item1;// grilla.DataTable.GetValue(0, filamatrix);
                        hoja.Cells[filacelda, 2] = v_item2;// grilla.DataTable.GetValue(1, filamatrix);
                        hoja.Cells[filacelda, 3] = v_item3;// grilla.DataTable.GetValue(2, filamatrix);
                        hoja.Cells[filacelda, 4] = v_item4;// grilla.DataTable.GetValue(3, filamatrix);
                        hoja.Cells[filacelda, 5] = v_item5;// grilla.DataTable.GetValue(4, filamatrix);
                        hoja.Cells[filacelda, 6] = v_item6;// grilla.DataTable.GetValue(5, filamatrix);
                        hoja.Cells[filacelda, 7] = v_item7;// grilla.DataTable.GetValue(6, filamatrix);
                        hoja.Cells[filacelda, 8] = v_item8;// grilla.DataTable.GetValue(7, filamatrix);
                        hoja.Cells[filacelda, 9] = v_item9;// grilla.DataTable.GetValue(8, filamatrix);
                        hoja.Cells[filacelda, 10] = v_item10;// grilla.DataTable.GetValue(9, filamatrix);

                        filacelda = filacelda + 1;
                        filamatrix = filamatrix + 1;
                        fila = fila + 1;
                    }
                    //creamos una carpeta en el escritorio para guardar el excel
                    string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string CarpEscr = path + "\\RUTEO";
                    if (!Directory.Exists(CarpEscr))
                    {
                        Directory.CreateDirectory(CarpEscr);
                    }
                    string v_texto = v_tipo.Selected.Description.ToString();
                    aplicacion.Visible = false;
                    aplicacion.UserControl = false;
                    string archivo = CarpEscr + "\\RUTEO_" + v_texto + "_" + DateTime.Now.Hour.ToString("D2") + "" + DateTime.Now.Minute.ToString("D2") + "" + DateTime.Now.Second.ToString("D2") + ".xls";
                    libro.SaveAs(archivo);
                    libro.Close();
                    aplicacion.Quit();
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Exportado con éxito", 1, "Ok");

                }
                catch (Exception e)
                {

                }
            }

        }
    }
}
