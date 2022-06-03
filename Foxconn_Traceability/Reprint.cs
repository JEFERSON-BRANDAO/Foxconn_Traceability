using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Windows;
using Classes;

namespace Foxconn_Traceability
{
    public partial class Reprint : Form
    {
        internal string SN_ = null;
        //
        public Reprint()
        {
            InitializeComponent();
            //
            #region RODAPÉ

            int anoCriacao = 2018;
            int anoAtual = DateTime.Now.Year;
            string texto = anoCriacao == anoAtual ? " Foxconn CNSBG All Rights Reserved." : "-" + anoAtual + " Foxconn CNSBG All Rights Reserved.";
            //
            lbRodape.Text = "Copyright © " + anoCriacao + texto;

            #endregion
            //
            this.txtSerial.Select();

            lbWo.Visible = false;
            txtWo.Visible = false;
        }

        private void GerarEtiqueta(string SN)
        {
            ARRIS_SN objSN = new ARRIS_SN();//ARRIS
            RokuSN objRokuSN = new RokuSN();//ROKU
            string serial = string.Empty;

            string modelo = objRokuSN.Modelo(SN);
            if (modelo.Contains("ROKU"))
            {
                if (SN.StartsWith("PM"))
                {
                    //ROKU-PANEL
                    string newSN = SN;
                    //serial = objRokuSN.pref_Roku_SN(modelo);
                    //
                    if (modelo.Contains("ERRO: ORA-12560: TNS:protocol adapter error"))
                    {
                        //para emitir som de alerta
                        Som objSom = new Som();
                        objSom.Falha();
                        //
                        MessageBox.Show(modelo + ". Verifique sua conexão de rede!", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        txtSerial.SelectAll();
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(modelo))
                        {
                            string lastSN = objRokuSN.get_lastSN(modelo);
                            //
                            if (!string.IsNullOrEmpty(newSN))
                            {
                                if (newSN.Length == lastSN.Length)
                                {
                                    if (string.IsNullOrEmpty(lastSN))
                                    {
                                        //para emitir som de alerta
                                        Som objSom = new Som();
                                        objSom.Falha();
                                        //
                                        MessageBox.Show("Nenhum serial encontrado", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        txtSerial.SelectAll();
                                    }
                                    else
                                    {
                                        if (lastSN.Contains("ERRO"))
                                        {
                                            //para emitir som de alerta
                                            Som objSom = new Som();
                                            objSom.Falha();
                                            //
                                            MessageBox.Show(lastSN, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            txtSerial.SelectAll();
                                        }
                                        else
                                        {
                                            #region REIMPRIMIR...

                                            Reimprimir(modelo, string.Empty);

                                            //txtSerial.SelectAll();
                                            txtSerial.Clear();
                                            txtSerial.Select();

                                            #endregion

                                        }
                                    }

                                }
                                else
                                {
                                    //para emitir som de alerta
                                    Som objSom = new Som();
                                    objSom.Falha();
                                    //
                                    MessageBox.Show("Tamanho do serial inválido. Tamanho SN = " + newSN.Length, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    txtSerial.SelectAll();
                                }
                            }
                            else
                            {
                                //para emitir som de alerta
                                Som objSom = new Som();
                                objSom.Falha();
                                //
                                MessageBox.Show("Serial não pode ser vázio", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                txtSerial.Select();
                            }

                        }
                        else
                        {
                            //para emitir som de alerta
                            Som objSom = new Som();
                            objSom.Falha();
                            //
                            MessageBox.Show("Não existe registro deste modelo na tabela R_AP_TEMP", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            txtSerial.SelectAll();
                        }
                    }
                }
                else if (modelo == "ROKU-MAC")
                {
                    //ROKU-MAC 
                    RokuSN etiqueta = new RokuSN();
                    serial = etiqueta.Get_MAC_WO(SN).Count > 0 ? etiqueta.Get_MAC_WO(SN)[0].MAC : string.Empty;

                    //serial = objRokuSN.Mac(SN);
                    if (!string.IsNullOrEmpty(serial))
                    {
                        string wo = etiqueta.Get_MAC_WO(SN)[0].WO;//txtSnVertical.Text.ToUpper().Trim();
                        //
                        if (wo.Length == 12)
                        {
                            #region REIMPRIMIR...

                            if (serial.Length == 12)
                            {
                                //string mac = string.Empty;
                                ////
                                //for (int linha = 0; linha < serial.Length; linha++)
                                //{
                                //    mac += serial[linha];//XX:XX:XX:XX:XX:XX
                                //    //
                                //    //if ((linha == 1) || (linha == 3) || (linha == 5) || (linha == 7) || (linha == 9))
                                //    //{
                                //    //    mac += ":";//adiciona : nas posições acima 
                                //    //}
                                //}
                                ////
                                //txtSerial.Text = mac;

                                if (wo.StartsWith("0000"))
                                {
                                    string wo_ = wo.Remove(0, 4);
                                    wo = wo_;
                                }
                                else if (wo.StartsWith("00"))
                                {
                                    string wo_ = wo.Remove(0, 2);
                                    wo = wo_;
                                }

                                Reimprimir("ROKU-MAC", wo);
                            }
                            else
                            {
                                //para emitir som de alerta
                                Som objSom = new Som();
                                objSom.Falha();
                                //
                                MessageBox.Show("Tamanho SN inválido " + serial, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }

                            //txtSerial.SelectAll();
                            txtSerial.Clear();
                            txtSerial.Select();

                            #endregion
                        }
                        else
                        {
                            //para emitir som de alerta
                            Som objSom = new Som();
                            objSom.Falha();
                            //
                            MessageBox.Show("Tamanho WO inválida " + wo, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else 
                    {
                        //para emitir som de alerta
                        Som objSom = new Som();
                        objSom.Falha();
                        //
                        MessageBox.Show("Serial inválido ou não registrado", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        txtSerial.SelectAll();
                    }
                }
                else if (modelo == "ROKU-SMT")
                {
                    //ROKU-SMT
                    string newSN = SN;
                    //serial = objRokuSN.pref_Roku_SN(modelo);
                    //
                    if (modelo.Contains("ERRO: ORA-12560: TNS:protocol adapter error"))
                    {
                        //para emitir som de alerta
                        Som objSom = new Som();
                        objSom.Falha();
                        //
                        MessageBox.Show(modelo + ". Verifique sua conexão de rede!", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        txtSerial.SelectAll();
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(modelo))
                        {
                            string lastSN = objRokuSN.get_lastSN(modelo);
                            //
                            if (!string.IsNullOrEmpty(newSN))
                            {
                                if (newSN.Length == lastSN.Length)
                                {
                                    if (string.IsNullOrEmpty(lastSN))
                                    {
                                        //para emitir som de alerta
                                        Som objSom = new Som();
                                        objSom.Falha();
                                        //
                                        MessageBox.Show("Nenhum serial encontrado", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        txtSerial.SelectAll();
                                    }
                                    else
                                    {
                                        if (lastSN.Contains("ERRO"))
                                        {
                                            //para emitir som de alerta
                                            Som objSom = new Som();
                                            objSom.Falha();
                                            //
                                            MessageBox.Show(lastSN, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            txtSerial.SelectAll();
                                        }
                                        else
                                        {
                                            #region REIMPRIMIR...

                                            //Reimprimir(modelo,string.Empty);

                                            RokuSN etiqueta = new RokuSN();
                                            string wo = etiqueta.Get_MAC_WO(SN)[0].WO;//txtSnVertical.Text.ToUpper().Trim();
                                            //
                                            if (wo.Length == 12)
                                            {
                                                #region REIMPRIMIR...

                                                //if (wo.StartsWith("0000"))
                                                //{
                                                //    string wo_ = wo.Remove(0, 4);
                                                //    wo = wo_;
                                                //}
                                                //else if (wo.StartsWith("00"))
                                                //{
                                                //    string wo_ = wo.Remove(0, 2);
                                                //    wo = wo_;
                                                //}

                                                Reimprimir(modelo, wo);

                                                //txtSerial.SelectAll();
                                                txtSerial.Clear();
                                                txtSerial.Select();

                                                #endregion
                                            }
                                            else
                                            {
                                                //para emitir som de alerta
                                                Som objSom = new Som();
                                                objSom.Falha();
                                                //
                                                MessageBox.Show("Tamanho WO inválida " + wo, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            }

                                            //txtSerial.SelectAll();
                                            txtSerial.Clear();
                                            txtSerial.Select();

                                            #endregion

                                        }
                                    }

                                }
                                else
                                {
                                    //para emitir som de alerta
                                    Som objSom = new Som();
                                    objSom.Falha();
                                    //
                                    MessageBox.Show("Tamanho do serial inválido. Tamanho SN = " + newSN.Length, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    txtSerial.SelectAll();
                                }
                            }
                            else
                            {
                                //para emitir som de alerta
                                Som objSom = new Som();
                                objSom.Falha();
                                //
                                MessageBox.Show("Serial não pode ser vázio", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                txtSerial.Select();
                            }
                        }
                        else
                        {
                            //para emitir som de alerta
                            Som objSom = new Som();
                            objSom.Falha();
                            //
                            MessageBox.Show("Não existe registro deste modelo na tabela R_AP_TEMP", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            txtSerial.SelectAll();
                        }
                    }
                }
            }
            else
            {
                //OUTROS MODELOS
                string newSN = SN;
                string wo = string.Empty;
                if (chkTg3442.Checked)
                {
                    string Z = "";
                    string A1 = "";
                    string B = "";
                    string A = "";
                    //
                    if (newSN.Length == 15)//Z075BA204500006
                    {
                        Z = newSN[0].ToString();
                        A1 = newSN[1].ToString();
                        B = newSN[4].ToString();
                        A = newSN[5].ToString();
                        //
                        if (A1 != "A")//Para não permitir TG1692 ou TG1692BR
                            newSN = Z + B + A;
                    }
                    //
                    if (newSN.Contains("ZBA"))
                    {
                        newSN = SN;
                        wo = Wo(newSN);
                        if (string.IsNullOrEmpty(wo))
                            wo = txtWo.Text.ToUpper().Trim();
                        //
                        if (string.IsNullOrEmpty(wo))
                        {
                            lbWo.Visible = true;
                            txtWo.Visible = true;
                            txtWo.SelectAll();

                            //para emitir som de alerta
                            Som objSom = new Som();
                            objSom.Falha();
                            MessageBox.Show("Informe número da WO", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        else
                        {
                            if (wo.Length < 12)
                            {
                                Som objSom = new Som();
                                objSom.Falha();
                                MessageBox.Show("Tamanho da WO inválido", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                            else
                            {
                                //4 últimos digitos da wo
                                string wo_Last4digits = "";
                                string aux = wo;
                                for (int index = 4; index >= 1; index--)
                                {
                                    wo_Last4digits += aux[aux.Length - index];
                                }
                                //
                                wo = wo_Last4digits;
                            }
                        }
                        //
                        modelo = "TG3442-SMT";
                    }
                    else
                    {
                        Som objSom = new Som();
                        objSom.Falha();
                        MessageBox.Show("Serial inválido", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                {
                    modelo = objSN.Modelo(newSN);
                }

                //
                if (modelo.Contains("ERRO: ORA-12560: TNS:protocol adapter error"))
                {
                    //para emitir som de alerta
                    Som objSom = new Som();
                    objSom.Falha();
                    //
                    MessageBox.Show(modelo + ". Verifique sua conexão de rede!", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtSerial.SelectAll();
                }
                else
                {
                    if (!string.IsNullOrEmpty(modelo))
                    {
                        string lastSN = objSN.get_lastSN(modelo);
                        //
                        if (!string.IsNullOrEmpty(newSN))
                        {
                            if (newSN.Length == lastSN.Length)
                            {
                                if (string.IsNullOrEmpty(lastSN))
                                {
                                    //para emitir som de alerta
                                    Som objSom = new Som();
                                    objSom.Falha();
                                    //
                                    MessageBox.Show("Nenhum serial encontrado", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    txtSerial.SelectAll();
                                }
                                else
                                {
                                    if (lastSN.Contains("ERRO"))
                                    {
                                        //para emitir som de alerta
                                        Som objSom = new Som();
                                        objSom.Falha();
                                        //
                                        MessageBox.Show(lastSN, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        txtSerial.SelectAll();
                                    }
                                    else
                                    {
                                        #region REIMPRIMIR...

                                        Reimprimir(modelo, wo);

                                        //txtSerial.SelectAll();
                                        txtSerial.Clear();
                                        txtSerial.Select();

                                        #endregion

                                        ////permite reimprimir somente se serial for menor ou igual ao último registro 
                                        //int ret = newSN.CompareTo(lastSN);
                                        ////
                                        //if (ret <= 0)// se -1[serial é menor];  se 0[serial é igual]
                                        //{
                                        //    #region REIMPRIMIR ...

                                        //    Reimprimir(modelo);                                      

                                        //    //txtSerial.SelectAll();
                                        //    txtSerial.Clear();
                                        //    txtSerial.Select();

                                        //    #endregion
                                        //}
                                        //else
                                        //{
                                        //    //para emitir som de alerta
                                        //    Som objSom = new Som();
                                        //    objSom.Falha();
                                        //    //
                                        //    MessageBox.Show("Permite reimprimir somente se serial for menor ou igual ao último registro", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        //    txtSerial.SelectAll();
                                        //}
                                    }
                                }

                            }
                            else
                            {
                                //para emitir som de alerta
                                Som objSom = new Som();
                                objSom.Falha();
                                //
                                MessageBox.Show("Tamanho do serial inválido. Tamanho SN = " + newSN.Length, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                txtSerial.SelectAll();
                            }
                        }
                        else
                        {
                            //para emitir som de alerta
                            Som objSom = new Som();
                            objSom.Falha();
                            //
                            MessageBox.Show("Serial não pode ser vázio", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            txtSerial.Select();
                        }

                    }
                    else
                    {
                        //para emitir som de alerta
                        Som objSom = new Som();
                        objSom.Falha();
                        //
                        MessageBox.Show("Não existe registro deste modelo na tabela R_AP_TEMP", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        txtSerial.SelectAll();
                    }
                }
            }

            if (txtWo.Visible)
            {
                txtWo.Text = string.Empty;
                txtWo.Visible = false;
                lbWo.Visible = false;
            }
            //
            SN_ = null;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void Reimprimir(string modelo, string wo)
        {
            #region REIMPRIMIR

            try
            {
                List<String> Lista_Serial = new List<String>();
                bool selecionarImpressora = chkSelecionar.Checked;
                //para emitir som de alerta
                Som objSom = new Som();
                objSom.Aprovado();
                //
                if ((modelo.Equals("TG3442-SMT") || modelo.Equals("TG1692BR-SMT") || (modelo.Equals("TG1692A-SMT")) || (modelo.Equals("E965")) || (modelo.Equals("DSI")) || (modelo.Equals("ROKU-MAC")) || (modelo.Equals("ROKU-PANEL")) || (modelo.Equals("ROKU-SMT"))))
                {
                    #region LABEL CODESOFT

                    PrintDialog pd = new PrintDialog();
                    string serial = txtSerial.Text.ToUpper().Trim();
                    string impressora = string.Empty;
                    string snHorizontal = string.Empty;
                    //
                    if (!selecionarImpressora)
                    {
                        pd.AllowSelection = true;
                        if (pd.ShowDialog() == DialogResult.OK)//exibe janela para selecionar impressora
                        {
                            impressora = pd.PrinterSettings.PrinterName;//nome da impressora selecionada
                        }
                        else
                        {
                            pd.AllowSelection = false;
                            MessageBox.Show("AVISO: Reimpressão cancelada", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return;
                        }
                    }
                    else
                    {
                        impressora = pd.PrinterSettings.PrinterName;//impressora padrão
                    }

                    if (modelo.Equals("ROKU-MAC"))
                        snHorizontal = "3930BR";

                    //LABEL codesoft                  
                    Lista_Serial.Add(serial);
                    Print codesoft = new Print();
                    codesoft.Etiqueta_CodeSoft(Lista_Serial, modelo, 1, 1, false, impressora, snHorizontal, wo, wo);
                    //LOG
                    Log objLog = new Log();
                    objLog.Gravar(serial, modelo, "Reimprimir");

                    #endregion
                }
                else
                {
                    #region LABEL ZPL

                    PrintDialog pd = new PrintDialog();
                    string serial = txtSerial.Text.ToUpper().Trim();
                    string impressora = string.Empty;
                    //
                    if (!selecionarImpressora)
                    {
                        pd.AllowSelection = true;
                        if (pd.ShowDialog() == DialogResult.OK)//exibe janela para selecionar impressora
                        {
                            impressora = pd.PrinterSettings.PrinterName;//nome da impressora selecionada
                        }
                        else
                        {
                            pd.AllowSelection = false;
                            MessageBox.Show("AVISO: Reimpressão cancelada", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return;
                        }
                    }
                    else
                    {
                        impressora = pd.PrinterSettings.PrinterName;//impressora padrão
                    }
                    //LABEL ZPL
                    Print zpl = new Print();
                    zpl.Reimprimir(serial, modelo, selecionarImpressora, impressora);
                    //LOG
                    Log objLog = new Log();
                    objLog.Gravar(serial, modelo, "Reimprimir");

                    #endregion
                }

            }
            catch (Exception ex)
            {
                Som objSom = new Som();
                objSom.Falha();
                //
                MessageBox.Show(ex.Message.ToString());
                //
                return;
            }

            #endregion
        }

        private string Wo(string sn)
        {
            OleDbConnect Objconn = new OleDbConnect();
            string workorder = string.Empty;
            //
            try
            {
                Objconn.Conectar();
                Objconn.Parametros.Clear();
                string sql = @"SELECT WORKORDERNO FROM mfworkstatus WHERE SYSSERIALNO='" + sn + "'";
                //                                  
                Objconn.SetarSQL(sql);
                Objconn.Executar();
            }
            finally
            {
                Objconn.Desconectar();
            }
            //
            if (Objconn.Tabela.Rows.Count > 0)
            {
                string wo = Objconn.Tabela.Rows[0]["WORKORDERNO"].ToString();
                //
                if (!string.IsNullOrEmpty(wo))
                {
                    workorder = wo;
                }
            }
            //
            return workorder;
        }

        private void txtSerial_KeyPress(object sender, KeyPressEventArgs e)
        {
            //Quando for usar scanner
            if (e.KeyChar != 13)
                SN_ += e.KeyChar.ToString().Replace("\r\n", string.Empty).Replace("\b", string.Empty).ToUpper().Trim();

            //
            if ((e.KeyChar == 13) || (e.KeyChar == Convert.ToChar(Keys.Enter)))
            {
                //caso valor digitado contenha símbolos
                if (SN_ != null)
                    SN_ = txtSerial.Text.Replace("\r\n", string.Empty).Replace("\b", string.Empty).ToUpper().Trim();
                //
                GerarEtiqueta(SN_);

                ////refresh
                //SN_ = txtSerial.Text.Replace("\r\n", string.Empty).Replace("\b", string.Empty).ToUpper().Trim();

            }
        }

    }
}
