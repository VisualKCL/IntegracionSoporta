using Microsoft.VisualBasic;
using SAPbobsCOM;
using System;
using System.Data.Odbc;
using System.Drawing.Printing;
using System.Globalization;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;

namespace IntegracionSoporta
{
    public partial class Form1 : Form
    {
        Funciones func = new Funciones();
        public Credenciales credenciales;
        private CultureInfo _nf = new System.Globalization.CultureInfo("en-US");
        public Form1()
        {
            InitializeComponent();
        }

        private void btCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btProcesar_Click(object sender, EventArgs e)
        {
            OdbcCommand cmdSql;
            OdbcDataAdapter dataAdapterSql;
            OdbcConnection connSql = null;
            System.Data.DataTable dtEncabezado = null;
            System.Data.DataTable dtDetalle = null;
            SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();
            SAPbobsCOM.Recordset ors;
            SAPbobsCOM.Documents odoc = null;
            String s;
            Int32 lRetCode = 0;
            String sErrMsg = "";

            try
            {
                lbLista.Items.Clear();
                credenciales = new Credenciales();
                func.ObtenerCredenciales(ref credenciales);

                if (func.ConectarSAP(ref credenciales, ref oCompany, ref lbLista))
                {
                    ors = ((SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset));
                    //buscar si documentos ya poseen folio
                    ActualizarFolios(ref oCompany);



                    if (func.ConectarSoporta(ref credenciales, ref lbLista, ref connSql))
                    {
                        try
                        {
                            connSql.Open();
                            s = @"SELECT TOP {1} CLIENTE
                                        ,TIPO_DOCUMENTO
                                        ,NRO_INT
                                        ,NRO_SII
                                        ,TIPO_DOCUMENTO
                                        ,MARCA_COMERCIALIZACION
                                        ,FECHA_DOCTO
                                        ,NRO_CONTRATO
                                        ,ID_CONTRATO_COCHILCO
                                        ,QUOTA
                                        ,PUERTO_EMBARQUE
                                        ,PUERTO_DESTINO
                                        ,NAVE
                                        ,LINEA_NAVIERA
                                        ,FORMA_PAGO
                                        ,BANCO
                                        ,ACREDITIVO
                                        ,MARCA
                                        ,ATADOS
                                        ,CATODOS
                                        ,PAIS_ORIGEN
                                        ,FECHA_BL
                                        ,NRO_BL
                                        ,MOD_ENTREGA
                                        ,PARTIDA_ARANCEL
                                        ,NIF
                                        ,PESO_NETO
                                        ,PESO_BRUTO
                                        ,TIPO_MONEDA
                                        ,EMPRESA
                                        ,IND_ESTADO
                                    FROM SOFTLAND_CABECERA
                                   WHERE 1 = 1
                                     AND FECHA_DOCTO >= '{0}'
                                   ORDER BY NRO_INT DESC ";
                            s = String.Format(s, credenciales.fechaInicio.ToString("yyyyMMdd"), credenciales.top);
                            //func.AddLog(s);
                            dataAdapterSql = new OdbcDataAdapter(s, connSql);
                            dtEncabezado = new System.Data.DataTable();
                            dataAdapterSql.Fill(dtEncabezado);
                            if (dtEncabezado.Rows.Count > 0)
                            {
                                var clave = 0;
                                foreach (System.Data.DataRow orow in dtEncabezado.Rows)
                                {
                                    var Mensaje = "";
                                    var nroInterno = "";
                                    var estado = "E";
                                    var docEntry = 0;
                                    var docNum = 0;
                                    var cliente = "";
                                    var docOrigen = "";
                                    var documento = "";
                                    var ObjType = "";
                                    var DocSubType = "";
                                    var CrearRegistoUDO = true;
                                    var tablaSAP = "";
                                    try
                                    {
                                        cliente = orow["CLIENTE"].ToString().Trim();
                                        var Origen = orow["TIPO_DOCUMENTO"].ToString().Trim();
                                        nroInterno = orow["NRO_INT"].ToString().Trim();
                                        if (Origen == "F")//si es Factura
                                        {
                                            odoc = ((SAPbobsCOM.Documents)oCompany.GetBusinessObject(BoObjectTypes.oOrders));//orden de venta
                                            documento = "Orden de Venta";
                                            docOrigen = "Factura";
                                            ObjType = "17";
                                            DocSubType = "--";
                                            tablaSAP = "ORDR";
                                        }
                                        else if (Origen == "C")//si es nota de credito
                                        {
                                            odoc = ((SAPbobsCOM.Documents)oCompany.GetBusinessObject(BoObjectTypes.oReturnRequest));//Solicitud de devolucion
                                            documento = "Solicituad Devolucion";
                                            docOrigen = "Nota Credito";
                                            ObjType = "234000031";
                                            DocSubType = "--";
                                            tablaSAP = "ORRR";
                                        }
                                        // si es nota de debito valor D
                                        else
                                        {
                                            documento = "Nota Debito preliminar";
                                            docOrigen = "Nota Debito";
                                            ObjType = "13";
                                            DocSubType = "DN";
                                            odoc = ((SAPbobsCOM.Documents)oCompany.GetBusinessObject(BoObjectTypes.oDrafts));//nota debito borrador
                                            odoc.DocObjectCode = BoObjectTypes.oInvoices;
                                            odoc.DocumentSubType = BoDocumentSubType.bod_DebitMemo;
                                            tablaSAP = "ODRF";
                                        }


                                        s = @"SELECT ""DocEntry"", ""DocStatus"" FROM ""{0}"" WHERE COALESCE(""U_VKS_NRO_INT"",'') = '{1}' ORDER BY ""CreateDate"" DESC, ""DocTime"" ";
                                        s = String.Format(s, tablaSAP, nroInterno);
                                        ors.DoQuery(s);
                                        if (ors.RecordCount > 0)
                                        {
                                            if ((Origen == "F") && (((String)ors.Fields.Item("DocStatus").Value).Trim() != "O"))//si es Factura
                                            {//busca si la orden de venta creada tiene factura y a su vez tiene nota de credito deja crear nuevamente en SAP
                                                s = @" SELECT COUNT(*) ""cant""
                                                        FROM ""ORDR"" T0
                                                        JOIN ""RDR1"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
                                                        JOIN ""OINV"" T2 ON T2.""DocEntry"" = T1.""TrgetEntry""
                                                                      AND T2.""ObjType"" = T1.""TargetType""
                                                        JOIN ""INV1"" T3 ON T3.""DocEntry"" = T2.""DocEntry""
                                                        JOIN ""ORIN"" T4 ON T4.""DocEntry"" = T3.""TrgetEntry""
                                                                      AND T4.""ObjType"" = T3.""TargetType""
                                                       WHERE T0.""DocEntry"" = {0}";
                                                s = String.Format(s, ((Int32)ors.Fields.Item("DocEntry").Value));
                                                /*
                                                 SELECT COUNT(*) "cant"
                                                        FROM "ORDR" T0
                                                        JOIN "RDR1" T1 ON T1."DocEntry" = T0."DocEntry"
                                                        JOIN "INV1" T3 ON T3."BaseEntry" = T0."DocEntry"
                                                                      AND T3."BaseType" = T0."ObjType"
                                                                      AND T3."BaseLine" = T1."LineNum"
                                                        JOIN "RIN1" T4 ON T4."BaseEntry" = T3."DocEntry"
                                                                      AND T4."BaseType" = T3."ObjType"
                                                                      AND T4."BaseLine" = T3."LineNum"
                                                       WHERE T0."DocEntry" = 121
                                                */
                                                ors.DoQuery(s);
                                                if (((Int32)ors.Fields.Item("cant").Value) == 0)
                                                {
                                                    CrearRegistoUDO = false;
                                                    continue;
                                                }
                                            }
                                            else
                                            {
                                                CrearRegistoUDO = false;
                                                continue;
                                            }
                                        }


                                        var nro_contrato = orow["NRO_CONTRATO"].ToString().Trim();
                                        s = @"SELECT ""CardCode"" FROM ""OCRD"" WHERE COALESCE(""AddID"",'') = '{0}' AND ""CardType"" = 'C' ";
                                        s = String.Format(s, nro_contrato.Substring(0, nro_contrato.IndexOf("-")));
                                        ors.DoQuery(s);
                                        if (ors.RecordCount == 0)
                                        {
                                            Mensaje = "Cliente no se ha encontrado en SAP";
                                            func.AddLog("Cliente no se ha encontrado en SAP: " + cliente + ". Nro Interno " + nroInterno);
                                            lbLista.Items.Add("Cliente no se ha encontrado en SAP: " + cliente + ". Nro Interno " + nroInterno);
                                            continue;
                                        }

                                        odoc.CardCode = ((String)ors.Fields.Item("CardCode").Value).Trim();
                                        odoc.DocDate = ((DateTime)orow["FECHA_DOCTO"]);
                                        odoc.DocDueDate = ((DateTime)orow["FECHA_DOCTO"]);
                                        odoc.TaxDate = ((DateTime)orow["FECHA_DOCTO"]);
                                        odoc.UserFields.Fields.Item("U_VKS_NRO_INT").Value = orow["NRO_INT"].ToString().Trim();
                                        //odoc.UserFields.Fields.Item("U_VKS_NRO_SII").Value = orow["NRO_SII"].ToString().Trim();
                                        odoc.UserFields.Fields.Item("U_VKS_TIPO_DOC").Value = orow["TIPO_DOCUMENTO"].ToString().Trim();
                                        odoc.UserFields.Fields.Item("U_VKS_MARCA_COM").Value = orow["MARCA_COMERCIALIZACION"].ToString().Trim();
                                        odoc.UserFields.Fields.Item("U_VKS_FECHA_DOC").Value = ((DateTime)orow["FECHA_DOCTO"]);//revisar

                                        odoc.UserFields.Fields.Item("U_FechaEmbarque").Value = ((DateTime)orow["FECHA_DOCTO"]);

                                        //odoc.UserFields.Fields.Item("U_VKS_NRO_CONTRATO").Value = orow["NRO_CONTRATO"].ToString().Trim();
                                        odoc.UserFields.Fields.Item("U_Contrato").Value = orow["NRO_CONTRATO"].ToString().Trim();

                                        //odoc.UserFields.Fields.Item("U_VKS_ID_CONT_COCHILCO").Value = orow["ID_CONTRATO_COCHILCO"].ToString().Trim();
                                        odoc.UserFields.Fields.Item("U_ConCoch").Value = orow["ID_CONTRATO_COCHILCO"].ToString().Trim();

                                        //odoc.UserFields.Fields.Item("U_VKS_QUOTA").Value = orow["QUOTA"].ToString().Trim();
                                        odoc.UserFields.Fields.Item("U_Quota").Value = orow["QUOTA"].ToString().Trim();


                                        var emba = orow["PUERTO_EMBARQUE"].ToString().Trim().Split(',');
                                        odoc.UserFields.Fields.Item("U_PaisEmb").Value = emba[1];
                                        //odoc.UserFields.Fields.Item("U_VKS_PTO_EMBARQUE").Value = emba[0];


                                        s = @"SELECT ""U_CodPuertoAduana"" FROM ""@VK_PTOSOPORTA"" WHERE ""Name"" = '{0}'";
                                        s = String.Format(s, emba[0]);
                                        ors.DoQuery(s);
                                        if (ors.RecordCount > 0)
                                            odoc.UserFields.Fields.Item("U_CodPtoEmbarque").Value = ((String)ors.Fields.Item("U_CodPuertoAduana").Value);
                                        else
                                        {
                                            s = @"SELECT ""Code""  FROM ""@VID_FEPTOEMB""  WHERE ""Name"" = '{0}'";
                                            s = String.Format(s, emba[0]);
                                            ors.DoQuery(s);
                                            if (ors.RecordCount > 0)
                                                odoc.UserFields.Fields.Item("U_CodPtoEmbarque").Value = ((String)ors.Fields.Item("Code").Value);
                                        }

                                        s = @"SELECT ""Code"" FROM ""@VID_FECODPAIS"" WHERE CASE WHEN COALESCE(""U_Descrip"", '') = '' THEN UPPER(""Name"") ELSE UPPER(""U_Descrip"") END = '{0}'";
                                        s = String.Format(s, emba[1].Trim().ToUpper());
                                        ors.DoQuery(s);
                                        if (ors.RecordCount > 0)
                                            odoc.UserFields.Fields.Item("U_CodPaisRecep").Value = ((String)ors.Fields.Item("Code").Value);



                                        var desti = orow["PUERTO_DESTINO"].ToString().Trim().Split(',');
                                        odoc.UserFields.Fields.Item("U_PuertoArribo").Value = desti[0];
                                        odoc.UserFields.Fields.Item("U_VKS_PtoDesemb").Value = desti[0];


                                        s = @"SELECT ""U_CodPuertoAduana"" FROM ""@VK_PTOSOPORTA"" WHERE ""Name"" = '{0}'";
                                        s = String.Format(s, desti[0]);
                                        ors.DoQuery(s);
                                        if (ors.RecordCount > 0)
                                            odoc.UserFields.Fields.Item("U_CodPtoDesemb").Value = ((String)ors.Fields.Item("U_CodPuertoAduana").Value);
                                        else
                                        {
                                            s = @"SELECT ""Code""  FROM ""@VID_FEPTOEMB""  WHERE ""Name"" = '{0}'";
                                            s = String.Format(s, desti[0]);
                                            ors.DoQuery(s);
                                            if (ors.RecordCount > 0)
                                                odoc.UserFields.Fields.Item("U_CodPtoDesemb").Value = ((String)ors.Fields.Item("Code").Value);
                                        }

                                        s = @"SELECT ""Code"" FROM ""@VID_FECODPAIS"" WHERE CASE WHEN COALESCE(""U_Descrip"", '') = '' THEN UPPER(""Name"") ELSE UPPER(""U_Descrip"") END = '{0}'";
                                        s = String.Format(s, desti[1].Trim().ToUpper());
                                        func.AddLog(s);
                                        ors.DoQuery(s);
                                        if (ors.RecordCount > 0)
                                            odoc.UserFields.Fields.Item("U_CodPaisDestin").Value = ((String)ors.Fields.Item("Code").Value);

                                        //odoc.UserFields.Fields.Item("U_VKS_NAVE").Value = orow["NAVE"].ToString().Trim();
                                        odoc.UserFields.Fields.Item("U_Nave").Value = orow["NAVE"].ToString().Trim();

                                        //odoc.UserFields.Fields.Item("U_VKS_LINEA_NAVIERA").Value = orow["LINEA_NAVIERA"].ToString().Trim();
                                        odoc.UserFields.Fields.Item("U_LineaNav").Value = orow["LINEA_NAVIERA"].ToString().Trim();

                                        //odoc.UserFields.Fields.Item("U_VKS_FORMA_PAGO").Value = orow["FORMA_PAGO"].ToString().Trim();
                                        odoc.UserFields.Fields.Item("U_Booking").Value = orow["FORMA_PAGO"].ToString().Trim();

                                        odoc.UserFields.Fields.Item("U_VKS_BANCO").Value = orow["BANCO"].ToString().Trim();
                                        odoc.UserFields.Fields.Item("U_VKS_ACREDITIVO").Value = orow["ACREDITIVO"].ToString().Trim();

                                        //odoc.UserFields.Fields.Item("U_VKS_MARCA").Value = orow["MARCA"].ToString().Trim();
                                        odoc.UserFields.Fields.Item("U_Marca").Value = orow["MARCA"].ToString().Trim();

                                        //odoc.UserFields.Fields.Item("U_VKS_ATADOS").Value = Convert.ToInt32(((Decimal)orow["ATADOS"]));//Numerico SAP
                                        odoc.UserFields.Fields.Item("U_Atados").Value = Convert.ToInt32(((Decimal)orow["ATADOS"])).ToString("N0").Replace(",", ".");//Numerico SAP

                                        //odoc.UserFields.Fields.Item("U_VKS_CATODOS").Value = Convert.ToInt32(((Decimal)orow["CATODOS"]));//Numerico SAP
                                        odoc.UserFields.Fields.Item("U_Catodos").Value = Convert.ToInt32(((Decimal)orow["CATODOS"])).ToString("N0").Replace(",", ".");

                                        //odoc.UserFields.Fields.Item("U_VKS_PAIS_ORIGEN").Value = orow["PAIS_ORIGEN"].ToString().Trim();
                                        odoc.UserFields.Fields.Item("U_PaisOrigen").Value = orow["PAIS_ORIGEN"].ToString().Trim();

                                        //odoc.UserFields.Fields.Item("U_VKS_FECHA_BL").Value = ((DateTime)orow["FECHA_BL"]);//revisar
                                        odoc.UserFields.Fields.Item("U_FechaBL").Value = ((DateTime)orow["FECHA_BL"]);

                                        //odoc.UserFields.Fields.Item("U_VKS_NRO_BL").Value = orow["NRO_BL"].ToString().Trim();
                                        odoc.UserFields.Fields.Item("U_NumeroBl").Value = orow["NRO_BL"].ToString().Trim();



                                        //odoc.UserFields.Fields.Item("U_VKS_MOD_ENTREGA").Value = orow["MOD_ENTREGA"].ToString().Trim();
                                        s = @"SELECT T1.""FldValue""
                                                  FROM ""CUFD"" T0
                                                  JOIN ""UFD1"" T1 ON T1.""TableID"" = T0.""TableID""
                                                                        AND T1.""FieldID"" = T0.""FieldID""
                                                 WHERE T0.""TableID"" = 'OINV'
                                                     AND T0.""AliasID"" = 'CodClauVenta'
                                                     AND T1.""Descr"" = '{0}'";
                                        s = String.Format(s, orow["MOD_ENTREGA"].ToString().Trim().Substring(0, 3));
                                        ors.DoQuery(s);
                                        if (ors.RecordCount > 0)
                                        {
                                            s = ((String)ors.Fields.Item("FldValue").Value);
                                            odoc.UserFields.Fields.Item("U_CodClauVenta").Value = s;
                                        }

                                        odoc.UserFields.Fields.Item("U_EntregaDelivery").Value = orow["MOD_ENTREGA"].ToString().Trim();
                                        odoc.UserFields.Fields.Item("U_TipoBultos").Value = "90";
                                        odoc.UserFields.Fields.Item("U_FmaPagExp").Value = "0";

                                        odoc.UserFields.Fields.Item("U_VK_PesoNeto").Value = ((Int32)orow["PESO_NETO"]).ToString("N0").Trim().Replace(",", ".") + " KGS";
                                        odoc.UserFields.Fields.Item("U_VK_PesoBruto").Value = ((Int32)orow["PESO_BRUTO"]).ToString("N0").Trim().Replace(",", ".") + " KGS";

                                        odoc.UserFields.Fields.Item("U_VKS_PARTIDA_ARC").Value = Convert.ToInt32(((Decimal)orow["PARTIDA_ARANCEL"]));//Numerico SAP
                                        odoc.UserFields.Fields.Item("U_VKS_NIF").Value = orow["NIF"].ToString().Trim();
                                        odoc.UserFields.Fields.Item("U_VKS_TIPO_MONEDA").Value = orow["TIPO_MONEDA"].ToString().Trim();
                                        odoc.UserFields.Fields.Item("U_VKS_EMPRESA").Value = orow["EMPRESA"].ToString().Trim();

                                        odoc.UserFields.Fields.Item("U_VKS_IND_ESTADO").Value = orow["IND_ESTADO"].ToString().Trim();
                                        odoc.UserFields.Fields.Item("U_VKS_SOPORTA").Value = "Y";
                                        odoc.UserFields.Fields.Item("U_CodModVenta").Value = "2";
                                        odoc.UserFields.Fields.Item("U_CodViaTransp").Value = "1";

                                        var sumaTotal = 0.0;
                                        //detalle
                                        s = @"SELECT LINEA
                                                    ,GLOSA
                                                    ,VALOR
                                                FROM SOFTLAND_ITEMS
                                               WHERE 1 = 1
                                                 AND NRO_INT = {0}";
                                        s = String.Format(s, orow["NRO_INT"].ToString().Trim());
                                        dataAdapterSql = new OdbcDataAdapter(s, connSql);
                                        dtDetalle = new System.Data.DataTable();
                                        dataAdapterSql.Fill(dtDetalle);
                                        var payto = "";
                                        var textopago = false;
                                        if (dtEncabezado.Rows.Count > 0)
                                        {
                                            var iLinea = 0;
                                            var sLinea = -1;
                                            var sLineas = 0;
                                            var errorDet = false;
                                            foreach (System.Data.DataRow orowD in dtDetalle.Rows)
                                            {
                                                try
                                                {


                                                    var texto = orowD["GLOSA"].ToString().Trim();
                                                    if (texto == "TOTAL:")
                                                    {
                                                        if (iLinea > 0)
                                                            odoc.Lines.Add();
                                                        odoc.Lines.ItemCode = credenciales.articulo;
                                                        odoc.Lines.ItemDescription = texto;
                                                        odoc.Lines.UnitPrice = Convert.ToDouble(orowD["VALOR"].ToString().Replace(",", ".").Trim(), _nf);
                                                        sumaTotal = sumaTotal + Convert.ToDouble(orowD["VALOR"].ToString().Replace(",", ".").Trim(), _nf);
                                                        odoc.Lines.TaxCode = "IVA_EXE";
                                                        iLinea++;
                                                        sLinea++;
                                                    }
                                                    else
                                                    {
                                                        if (texto == "PLEASE PAY TO:")
                                                        {
                                                            textopago = true;
                                                            continue;
                                                        }

                                                        if (textopago)
                                                        {
                                                            if (payto == "")
                                                                payto = texto;
                                                            else
                                                                payto = payto + Environment.NewLine + texto;
                                                        }
                                                        else
                                                        {
                                                            if (sLineas > 0)
                                                                odoc.SpecialLines.Add();

                                                            odoc.SpecialLines.LineText = texto;
                                                            odoc.SpecialLines.LineType = BoDocSpecialLineType.dslt_Text;
                                                            odoc.SpecialLines.AfterLineNumber = sLinea;
                                                            sLineas++;
                                                        }
                                                    }
                                                }
                                                catch (Exception dd)
                                                {
                                                    func.AddLog("Error Documento " + documento + " en detalle: " + dd.Message + (credenciales.mostrarLineError ? ", TRACE " + dd.StackTrace : ""));
                                                    lbLista.Items.Add("Error Documento " + documento + " en detalle: " + dd.Message);
                                                    Mensaje = @"Problema en detalle: " + dd.Message;
                                                    errorDet = true;
                                                    break;
                                                }
                                                finally
                                                {

                                                }
                                            }

                                            if (errorDet)
                                                continue;
                                        }
                                        else
                                        {
                                            func.AddLog("Documento " + documento + ": no se ha encontrado detalle");
                                            lbLista.Items.Add("Documento " + documento + ": no se ha encontrado detalle");
                                            Mensaje = "No se ha encontrado detalle";
                                            continue;
                                        }

                                        odoc.UserFields.Fields.Item("U_TotClauVenta").Value = sumaTotal;
                                        odoc.UserFields.Fields.Item("U_VKS_DatosBancarios").Value = payto;

                                        lRetCode = odoc.Add();
                                        if (lRetCode != 0)
                                        {
                                            String sPath = Path.GetDirectoryName(this.GetType().Assembly.Location);
                                            odoc.SaveXML(sPath + @"\\DOCs\\Documento_" + nroInterno.ToString());
                                            oCompany.GetLastError(out lRetCode, out sErrMsg);
                                            Mensaje = sErrMsg;
                                            func.AddLog("No se ha creado documento " + documento + ": " + sErrMsg);
                                            lbLista.Items.Add("No se ha creado documento " + documento + ": " + sErrMsg);
                                        }
                                        else
                                        {
                                            estado = "C";
                                            Mensaje = "Documento creado";
                                            docEntry = Convert.ToInt32(oCompany.GetNewObjectKey());
                                            s = @"SELECT ""DocNum"" FROM ""{0}"" WHERE ""DocEntry"" = {1}";
                                            s = String.Format(s, (Origen == "F" ? "ORDR" : (Origen == "C" ? "ORRR" : "ODRF")), docEntry);
                                            //func.AddLog(s);
                                            ors.DoQuery(s);
                                            docNum = ((Int32)ors.Fields.Item("DocNum").Value);
                                            func.AddLog("Documento " + documento + " creado: Nro SAP " + docNum.ToString());
                                            lbLista.Items.Add("Se ha creado documento " + documento + ": " + ors.Fields.Item("DocNum").Value.ToString());
                                        }

                                    }
                                    catch (Exception xx)
                                    {
                                        func.AddLog("Error Procesar documento " + nroInterno + ": " + xx.Message + (credenciales.mostrarLineError ? ", TRACE " + xx.StackTrace : ""));
                                        lbLista.Items.Add("Error Procesar documento " + nroInterno + ": " + xx.Message + (credenciales.mostrarLineError ? ", TRACE " + xx.StackTrace : ""));
                                        Mensaje = xx.Message;
                                    }
                                    finally
                                    {
                                        if (CrearRegistoUDO)
                                            AddUDORecord(ref oCompany, nroInterno, docOrigen, documento, cliente, Mensaje, estado, docEntry, docNum, ObjType, DocSubType);
                                    }
                                }
                            }
                        }
                        catch (Exception zx)
                        {
                            func.AddLog("Error Procesar: " + zx.Message + (credenciales.mostrarLineError ? ", TRACE " + zx.StackTrace : ""));
                            lbLista.Items.Add("Error Procesar: " + zx.Message + (credenciales.mostrarLineError ? ", TRACE " + zx.StackTrace : ""));
                        }
                        finally
                        {
                            connSql.Close();
                            lbLista.Items.Add("Desconectado de Soporta");
                        }
                    }

                    oCompany.Disconnect();
                    lbLista.Items.Add("Desconectado de SAP");
                }

            }
            catch (Exception ex)
            {
                func.AddLog("Error btProcesar: " + ex.Message + (credenciales.mostrarLineError ? ", TRACE " + ex.StackTrace : ""));
            }
        }


        public void AddUDORecord(ref SAPbobsCOM.Company oCompany, String Nro_Interno, String Tipo_Doc, String Tipo_Doc_Dest, String Cliente, String Mensaje, String Estado, Int32 DocEntry, Int32 DocNum, String ObjType, String DocSubType)
        {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralData oChild = null;
            SAPbobsCOM.GeneralDataCollection oChildren = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;
            SAPbobsCOM.Recordset oQuery = ((SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset));
            String query;
            try
            {
                query = @"SELECT COUNT(*) ""cant"" FROM ""@VID_SOPORTALOG"" WHERE ""U_Nro_Interno"" = '{0}' AND ""U_Estado"" <> 'C' ";
                query = String.Format(query, Nro_Interno);
                oQuery.DoQuery(query);

                if (((Int32)oQuery.Fields.Item("cant").Value) > 0)
                {
                    query = @"UPDATE ""@VID_SOPORTALOG"" SET ""U_Estado"" = '{1}', ""U_Mensaje"" = '{2}', ""U_DocEntry"" = {3}, ""U_DocNum"" = {4}
                               WHERE ""U_Nro_Interno"" = '{0}' ";
                    query = String.Format(query, Nro_Interno, Estado, Mensaje, DocEntry, DocNum);
                    oQuery.DoQuery(query);
                }
                else
                {
                    oCompanyService = oCompany.GetCompanyService();
                    // Get GeneralService (oCmpSrv is the CompanyService)
                    oGeneralService = oCompanyService.GetGeneralService("VID_SOPORTALOG");
                    // Create data for new row in main UDO
                    oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));
                    //oGeneralData.SetProperty("Code", Nro_Interno.PadLeft(12, '0')); //ya no va se paso a UDO Documento
                    oGeneralData.SetProperty("U_Nro_Interno", Nro_Interno);
                    oGeneralData.SetProperty("U_Tipo_Doc", Tipo_Doc);
                    oGeneralData.SetProperty("U_Tipo_Doc_Dest", Tipo_Doc_Dest);
                    oGeneralData.SetProperty("U_Cliente", Cliente);
                    oGeneralData.SetProperty("U_Mensaje", Mensaje);
                    oGeneralData.SetProperty("U_Estado", Estado);
                    oGeneralData.SetProperty("U_ObjType", ObjType);
                    oGeneralData.SetProperty("U_DocSubType", DocSubType);
                    //func.AddLog("DocEntry: " + DocEntry.ToString() + " DocNum: " + DocNum.ToString());
                    oGeneralData.SetProperty("U_DocEntry", DocEntry);
                    oGeneralData.SetProperty("U_DocNum", DocNum);
                    //  Handle child rows
                    //oChildren = oGeneralData.Child("SM_MOR1");
                    //int i = 0;
                    //for (i = 1; i <= lstMainDish.Items.Count; i++)
                    //{
                    //    // Create data for rows in the child table
                    //    oChild = oChildren.Add();
                    //    oChild.SetProperty("U_MainDish", lstMainDish.Items[i - 1]);
                    //    oChild.SetProperty("U_SideDish", lstSideDish.Items[i - 1]);
                    //    oChild.SetProperty("U_Drink", lstDrink.Items[i - 1]);
                    //}
                    // Add the new row, including children, to database
                    oGeneralParams = oGeneralService.Add(oGeneralData);
                    //txtCode.Text = System.Convert.ToString(oGeneralParams.GetProperty("DocEntry"));
                    //Interaction.MsgBox("Record added", (Microsoft.VisualBasic.MsgBoxStyle)(0), null);
                }
            }
            catch (Exception ex)
            {
                func.AddLog("Error AddUDORecord: " + ex.Message + (credenciales.mostrarLineError ? ", TRACE " + ex.StackTrace : ""));
            }
        }

        public void ActualizarFolios(ref SAPbobsCOM.Company oCompany)
        {
            SAPbobsCOM.Recordset oQuery = ((SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset));
            SAPbobsCOM.Recordset oQuery2 = ((SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset));
            String query;
            try
            {
                query = @"SELECT T0.""DocEntry""
                              ,CASE COALESCE(T0.""U_ObjType"",'')
			                        WHEN '13' THEN 'OINV'
			                        WHEN '234000031' THEN 'ORRR'
			                        ELSE 'ORDR'--Orden de venta
                               END ""tabla""
                              ,CASE COALESCE(T0.""U_ObjType"",'')
			                        WHEN '13' THEN 'INV1'
			                        WHEN '234000031' THEN 'RRR1'--solicitud de devolucion
			                        ELSE 'RDR1'--Orden de venta
                               END ""tablaDet""
                              ,CASE COALESCE(T0.""U_ObjType"",'')
			                        WHEN '13' THEN 'OINV'
			                        WHEN '234000031' THEN 'ORIN'--solicitud de devolucion
			                        ELSE 'OINV'--Orden de venta
                               END ""tablaDestino""
                              ,T0.""U_DocSubType""
                              ,T0.""U_DocEntry""
                              ,T0.""U_ObjType""
                          FROM ""@VID_SOPORTALOG"" T0
                         WHERE COALESCE(T0.""U_FolioNum"",0) = 0
                           AND T0.""U_Estado"" = 'C'
                           AND COALESCE(T0.""U_DocEntry"",0) > 0";
                oQuery.DoQuery(query);

                if (oQuery.RecordCount == 0)
                {
                    func.AddLog("No se han encontrado registros para actualizar folios");
                    lbLista.Items.Add("No se han encontrado registros para actualizar folios");
                }
                else
                {
                    while (!oQuery.EoF)
                    {
                        try
                        {
                            if ((((String)oQuery.Fields.Item("U_ObjType").Value).Trim() == "13") && (((String)oQuery.Fields.Item("U_DocSubType").Value).Trim() == "DN"))
                            {
                                query = @"SELECT T2.""FolioPref""
                                        ,T2.""FolioNum""
                                    FROM ""{0}"" T2
                                   WHERE COALESCE(T2.""draftKey"",0) = {1}
                                     AND COALESCE(T2.""FolioNum"",0) > 0";
                                query = String.Format(query, ((String)oQuery.Fields.Item("tabla").Value).Trim(), ((Int32)oQuery.Fields.Item("U_DocEntry").Value));
                            }
                            else
                            {
                                query = @"SELECT T2.""FolioPref""
                                        ,T2.""FolioNum""
                                    FROM ""{0}"" T0
                                    JOIN ""{1}"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
                                    JOIN ""{3}"" T2 ON T2.""DocEntry"" = T1.""TrgetEntry""
                                                  AND T2.""ObjType"" = T1.""TargetType""
                                   WHERE 1 = 1
                                     AND COALESCE(T2.""FolioNum"",0) > 0
                                     AND T0.""DocEntry"" = {2}";
                                query = String.Format(query, ((String)oQuery.Fields.Item("tabla").Value).Trim(), ((String)oQuery.Fields.Item("tablaDet").Value).Trim(), ((Int32)oQuery.Fields.Item("U_DocEntry").Value), ((String)oQuery.Fields.Item("tablaDestino").Value).Trim());
                            }
                            oQuery2.DoQuery(query);
                            if (oQuery2.RecordCount > 0)
                            {
                                var pref = ((String)oQuery2.Fields.Item("FolioPref").Value).Trim();
                                var fol = ((Int32)oQuery2.Fields.Item("FolioNum").Value);
                                query = @" UPDATE ""@VID_SOPORTALOG"" SET ""U_FolioPref"" = '{1}', ""U_FolioNum"" = {2} WHERE ""DocEntry"" = {0}";
                                query = String.Format(query, ((Int32)oQuery.Fields.Item("DocEntry").Value), pref, fol);
                                oQuery2.DoQuery(query);
                            }
                        }
                        catch (Exception xx)
                        {
                            func.AddLog("Error ActualizarFolios: actualizar folio DocEntry " + oQuery.Fields.Item("DocEntry").Value.ToString() + " ," + xx.Message + (credenciales.mostrarLineError ? ", TRACE " + xx.StackTrace : ""));
                        }
                        finally
                        {
                            oQuery.MoveNext();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                func.AddLog("Error ActualizarFolios: " + ex.Message + (credenciales.mostrarLineError ? ", TRACE " + ex.StackTrace : ""));
            }
            finally
            {
                oQuery = null;
                oQuery2 = null;
            }
        }

    }
}