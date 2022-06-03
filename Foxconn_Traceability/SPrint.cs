using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Threading;
using Classes;
using System.Configuration;
using System.Globalization;

namespace Foxconn_Traceability
{

    public partial class Form1 : Form
    {
        internal string mensagemStatus = "Processando... ";

        public Form1()
        {
            InitializeComponent();
            //
            Clear_ProgressBar();

            #region RODAPÉ

            int anoCriacao = 2018;
            int anoAtual = DateTime.Now.Year;
            string texto = anoCriacao == anoAtual ? " Foxconn CNSBG All Rights Reserved." : "-" + anoAtual + " Foxconn CNSBG All Rights Reserved.";
            //
            lbRodape.Text = "Copyright © " + anoCriacao + texto;

            #endregion
            //           
            Configuracao impressora = new Configuracao();
            if (impressora.Permite_Imprimir())
            {
                CarregaComboModelos();
                txtQty.Select();
                //
                btnGerar.Enabled = true;
                cboModelo.Enabled = true;
                txtQty.Enabled = true;

                txtValor.Enabled = true;
                chkEtiquetaDupla.Enabled = true;
                chkImpressoraPadrao.Enabled = true;
            }
            else
            {
                btnGerar.Enabled = false;
                cboModelo.Enabled = false;
                txtQty.Enabled = false;

                txtValor.Enabled = false;
                chkEtiquetaDupla.Enabled = false;
                chkImpressoraPadrao.Enabled = false;

            }

        }

        public void CarregaComboModelos()
        {
            string caminho = AppDomain.CurrentDomain.BaseDirectory + @"\CONFIGURACAO\MODELO.txt";
            string linha;
            //
            cboModelo.Items.Add("----SELECIONE----");
            cboModelo.SelectedItem = "----SELECIONE----";
            //
            if (System.IO.File.Exists(caminho))
            {
                System.IO.StreamReader arqTXT = new System.IO.StreamReader(caminho);
                //
                while ((linha = arqTXT.ReadLine()) != null)
                {
                    if (!string.IsNullOrEmpty(linha))
                    {
                        if (!cboModelo.Items.Contains(linha.ToUpper().Trim()))//não add duplicado
                            cboModelo.Items.Add(linha.ToUpper().Trim());
                    }
                }

                //deixa selcionado item 0
                cboModelo.SelectedIndex = 0;
                cboModelo.SelectAll();
            }

        }

        public void createNewRecords(int valor, int qtdMaxima)
        {
            progressBar1.Maximum = qtdMaxima;
            progressBar1.Value += valor;
            // Sets the progress bar's Maximum property to  
            // the total number of records to be created.  
            //progressBar1.Maximum = 20;

            // Creates a new record in the dataset.  
            // NOTE: The code below will not compile, it merely  
            // illustrates how the progress bar would be used. 
            // Increases the value displayed by the progress bar.  
            //progressBar1.Value += 1;
            // Updates the label to show that a record was read. 

            if (progressBar1.Value == qtdMaxima)
            {
                lbProcessando.Text = "Concluído " + progressBar1.Value.ToString();
            }
            else
            {
                lbProcessando.Text = "De: " + progressBar1.Value.ToString() + " Até: " + qtdMaxima;//mensagemStatus + progressBar1.Value.ToString() + "%";
            }
        }

        public void Clear_ProgressBar()
        {
            progressBar1.Value = 0;
            lbProcessando.Text = string.Empty;//"Records Read = " + progressBar1.Value.ToString();
        }

        private void btnGerar_Click(object sender, EventArgs e)
        {
            Clear_ProgressBar();
            //
            string quantidade = txtQty.Text.Trim();
            string rowItem = cboModelo.SelectedIndex.ToString();
            //
            if (rowItem != "0")
            {
                if (string.IsNullOrEmpty(quantidade))
                {
                    MessageBox.Show("Quantidade não pode ser vázia", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //para emitir som de alerta
                    Som objSom = new Som();
                    objSom.Falha();
                }
                else
                {
                    if (int.Parse(quantidade) > 0)
                    {
                        #region Imprimir....

                        Imprimir();

                        //#region EXIBE progressBar ....

                        //bgWorkerIndeterminada.RunWorkerAsync();

                        //#endregion

                        //
                        txtQty.Select();

                        #endregion

                    }
                    else
                    {
                        MessageBox.Show("Quantide dever ser maior que 0", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        //para emitir som de alerta
                        Som objSom = new Som();
                        objSom.Falha();
                    }
                }
            }
            else
            {
                MessageBox.Show("Selecione um Modelo", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //para emitir som de alerta
                Som objSom = new Som();
                objSom.Falha();
            }

        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja fechar a aplicação?", "Fechar Aplicação",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    == DialogResult.Yes)
            {
                Application.Exit();
            }

        }

        private void btnReprint_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            //
            Reprint reprint = new Reprint();
            reprint.ShowDialog();
            //
            this.Visible = true;
        }

        public void Imprimir()
        {
            try
            {
                string modelo = cboModelo.SelectedItem.ToString().Trim();
                int quantidade = Int32.Parse(txtQty.Text.Trim());
                bool etiquetaDupla = chkEtiquetaDupla.Checked;
                bool UsarImpressoraPadrao = chkImpressoraPadrao.Checked;
                string impressora = string.Empty;

                PrintDialog pd = new PrintDialog();
                if (!UsarImpressoraPadrao)
                {
                    pd.AllowSelection = true;
                    if (pd.ShowDialog() == DialogResult.OK)
                    {
                        //exibe janela para selecionar impressora
                        impressora = pd.PrinterSettings.PrinterName;
                    }
                    else
                    {
                        Clear_ProgressBar();
                        pd.AllowSelection = false;
                        MessageBox.Show("AVISO: Impressão cancelada", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }
                else
                {
                    impressora = pd.PrinterSettings.PrinterName;
                }
                // 
                if ((modelo.Equals("TG3442-SMT")) || (modelo.Equals("TG1692A-SMT")) || (modelo.Equals("TG1692BR-SMT") || (modelo.Equals("ROKU-MAC")) || (modelo.Equals("ROKU-PANEL")) || (modelo.Equals("ROKU-SMT")) || (modelo.Equals("ROKU-GIFIT_BOX"))))
                {
                    string snHorizontal = string.Empty;
                    string snVertical = string.Empty;
                    string wo_Last3digits = string.Empty;
                    string wo_Last4digits = string.Empty;

                    if (modelo.Equals("ROKU-MAC"))
                    {
                        #region ROKU MAC

                        RokuSN SN = new RokuSN();
                        snHorizontal = "3930BR";//txtSnHorizontal.Text.ToUpper().Trim();
                        snVertical = txtValor.Text.ToUpper().Trim();
                        //
                        if (string.IsNullOrEmpty(snVertical))
                        {
                            //para emitir som de alerta
                            Som objSom = new Som();
                            objSom.Falha();
                            // 
                            Clear_ProgressBar();
                            MessageBox.Show("Sn Horizontal e Sn Vertical são obrigatórios", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        string qtd_wo = SN.get_Wo(snVertical);
                        if (int.Parse(qtd_wo) == 0)
                        {
                            //para emitir som de alerta
                            Som objSom = new Som();
                            objSom.Falha();
                            // 
                            Clear_ProgressBar();
                            MessageBox.Show("Wo inválida ou fechada ou cheia", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        if (int.Parse(qtd_wo) < quantidade)
                        {
                            //para emitir som de alerta
                            Som objSom = new Som();
                            objSom.Falha();
                            // 
                            Clear_ProgressBar();
                            MessageBox.Show("Quantidade informada é maior que quantidade da WO", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        //
                        if (snVertical.StartsWith("0000"))
                        {
                            string wo_ = snVertical.Remove(0, 4);
                            snVertical = wo_;
                        }
                        else if (snVertical.StartsWith("00"))
                        {
                            string wo_ = snVertical.Remove(0, 2);
                            snVertical = wo_;
                        }
                        //
                        if (etiquetaDupla)
                        {
                            if (quantidade % 2 != 0)
                            {
                                //para emitir som de alerta
                                Som objSom = new Som();
                                objSom.Falha();
                                // 
                                Clear_ProgressBar();
                                MessageBox.Show("Quantidade de Impressão somente números pares", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                            else
                            {
                                //RokuSN SN = new RokuSN();
                                List<String> Seriais = SN.get_Roku_Label(quantidade, modelo);
                                //
                                if (Seriais.Count <= 0)
                                {
                                    //para emitir som de alerta
                                    Som objSom = new Som();
                                    objSom.Falha();
                                    // 
                                    Clear_ProgressBar();
                                    MessageBox.Show("Erro ao gerar seriais", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return;
                                }

                                List<String> Lista_Serial_Duplo = new List<String>();
                                string ladoEsquerdo = string.Empty;
                                string ladoDireito = string.Empty;
                                //
                                for (int index = 0; index < quantidade; index++)
                                {
                                    if (index == 0)
                                    {
                                        ladoEsquerdo = Seriais[index].ToString();
                                    }
                                    else
                                    {
                                        if (index % 2 == 0)//par
                                            ladoEsquerdo = Seriais[index].ToString();
                                        if (index % 2 != 0)//impar
                                            ladoDireito = Seriais[index].ToString();
                                    }

                                    if ((ladoEsquerdo != string.Empty) && (ladoDireito != string.Empty))
                                    {
                                        string sn1_sn2 = ladoEsquerdo + ";" + ladoDireito;
                                        Lista_Serial_Duplo.Add(sn1_sn2);
                                        //
                                        ladoEsquerdo = string.Empty;
                                        ladoDireito = string.Empty;
                                    }

                                }
                                //
                                if (Lista_Serial_Duplo.Count > 0)
                                {
                                    int de = 1;
                                    int ate = Lista_Serial_Duplo.Count;
                                    //
                                    for (int index = 0; index < ate; index++)
                                    {
                                        //VALOR BARRA DE STATUS
                                        createNewRecords(de, ate);

                                        string serial = Lista_Serial_Duplo[index].ToString();
                                        List<String> sn = new List<String>();
                                        sn.Add(serial);

                                        //LABEL codesoft
                                        Print codesoft = new Print();
                                        codesoft.Etiqueta_CodeSoft(sn, modelo, 1, quantidade, etiquetaDupla, impressora, snHorizontal, snVertical, "");

                                        //LOG                                    
                                        Log objLog = new Log();
                                        objLog.Gravar(serial, modelo, "Imprimir");
                                    }
                                }

                            }

                        }
                        else
                        {
                            if (int.Parse(qtd_wo) == 0)
                            {
                                //para emitir som de alerta
                                Som objSom = new Som();
                                objSom.Falha();
                                // 
                                Clear_ProgressBar();
                                MessageBox.Show("Wo inválida ou fechada ou cheia", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }

                            if (int.Parse(qtd_wo) < quantidade)
                            {
                                //para emitir som de alerta
                                Som objSom = new Som();
                                objSom.Falha();
                                // 
                                Clear_ProgressBar();
                                MessageBox.Show("Quantidade informada é maior que quantidade da WO", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                            //RokuSN SN = new RokuSN();
                            List<String> Seriais = SN.get_Roku_Label(quantidade, modelo);
                            //
                            if (Seriais.Count <= 0)
                            {
                                //para emitir som de alerta
                                Som objSom = new Som();
                                objSom.Falha();
                                //   
                                Clear_ProgressBar();
                                MessageBox.Show("Erro ao gerar seriais", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                            //
                            if (Seriais.Count > 0)
                            {

                                int de = 1;
                                int ate = Seriais.Count;

                                for (int index = 0; index < ate; index++)
                                {
                                    //VALOR BARRA DE STATUS
                                    createNewRecords(de, ate);

                                    string serial = Seriais[index].ToString();
                                    List<String> Sn = new List<String>();
                                    Sn.Add(serial);

                                    //LABEL codesoft
                                    Print codesoft = new Print();
                                    codesoft.Etiqueta_CodeSoft(Sn, modelo, quantidade, quantidade, etiquetaDupla, impressora, snHorizontal, snVertical, "");

                                    //LOG
                                    Log objLog = new Log();
                                    objLog.Gravar(serial, modelo, "Imprimir");
                                }
                            }

                        }

                        #endregion
                    }
                    else if (modelo.Equals("ROKU-SMT"))//ROKU-BLANK
                    {
                        #region ROKU SMT

                        if (quantidade % 4 != 0)
                        {
                            Clear_ProgressBar();
                            MessageBox.Show("Quantidade deve ser múltiplo de 4", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }


                        RokuSN Wo = new RokuSN();
                        string workworder = txtValor.Text.Trim();
                        if (!Wo.Wo_Valida(workworder))
                        {
                            Clear_ProgressBar();
                            MessageBox.Show("Wo inválida", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }


                        if (etiquetaDupla)
                        {
                            if (quantidade % 2 != 0)
                            {
                                //para emitir som de alerta
                                Som objSom = new Som();
                                objSom.Falha();
                                // 
                                Clear_ProgressBar();
                                MessageBox.Show("Quantidade de Impressão somente números pares", "", MessageBoxButtons.OK, MessageBoxIcon.Error);

                                return;
                            }
                            else
                            {
                                RokuSN SN = new RokuSN();
                                List<String> Seriais = SN.get_Roku_Label(quantidade, modelo);
                                //
                                if (Seriais.Count <= 0)
                                {
                                    //para emitir som de alerta
                                    Som objSom = new Som();
                                    objSom.Falha();
                                    // 
                                    Clear_ProgressBar();
                                    MessageBox.Show("Erro ao gerar seriais", "", MessageBoxButtons.OK, MessageBoxIcon.Error);

                                    return;
                                }

                                List<String> Lista_Serial_Duplo = new List<String>();
                                string serial = string.Empty;
                                string serial2 = string.Empty;
                                string serial3 = string.Empty;
                                string serial4 = string.Empty;
                                //string wo = txtValor.Text.Trim().ToUpper();
                                //
                                for (int index = 0; index < quantidade; index++)
                                {
                                    if (index == 0)
                                    {
                                        serial = Seriais[index].ToString();
                                    }
                                    else
                                    {
                                        if (index % 2 == 0)
                                        {   //par
                                            serial = Seriais[index].ToString();
                                        }
                                        if (index % 2 != 0)
                                        {   //impar
                                            serial2 = Seriais[index].ToString();
                                        }
                                    }

                                    if ((serial != string.Empty) && (serial2 != string.Empty))
                                    {
                                        string sn1_sn2 = serial + ";" + serial2;
                                        Lista_Serial_Duplo.Add(sn1_sn2);
                                        //
                                        serial = string.Empty;
                                        serial2 = string.Empty;
                                    }

                                }
                                //
                                if (Lista_Serial_Duplo.Count > 0)
                                {
                                    int de = 1;
                                    int ate = Lista_Serial_Duplo.Count;
                                    // 
                                    List<String> listaSeriais = new List<String>();
                                    for (int index = 0; index < ate; index++)
                                    {
                                        //VALOR BARRA DE STATUS
                                        createNewRecords(de, ate);

                                        if ((index % 2 == 0) || index == 0)
                                        {
                                            serial = Lista_Serial_Duplo[index].ToString();
                                            List<String> sn = new List<String>();
                                            sn.Add(serial);
                                            listaSeriais = sn;
                                        }
                                        else
                                        {
                                            string serial_ = Lista_Serial_Duplo[index].ToString();
                                            if (serial_.Length == 15)
                                            {
                                                for (int linha = 0; linha < 7; linha++)
                                                {
                                                    serial3 += serial_[linha].ToString();
                                                }

                                                for (int linha = 8; linha < 15; linha++)
                                                {
                                                    serial4 += serial_[linha].ToString();
                                                }

                                                //LABEL codesoft
                                                Print codesoft = new Print();
                                                codesoft.Etiqueta_CodeSoft(listaSeriais, modelo, 1, quantidade, etiquetaDupla, impressora, serial3, serial4, workworder);

                                                //LOG                                    
                                                Log objLog = new Log();
                                                objLog.Gravar(serial, modelo, "Imprimir");

                                                serial3 = string.Empty;
                                                serial4 = string.Empty;
                                            }
                                            //else
                                            //{
                                            //    //erro
                                            //}
                                        }


                                        ////LABEL codesoft
                                        //Print codesoft = new Print();
                                        //codesoft.Etiqueta_CodeSoft(sn, modelo, 1, quantidade, etiquetaDupla, impressora, serial3, serial4, wo);

                                        ////LOG                                    
                                        //Log objLog = new Log();
                                        //objLog.Gravar(serial, modelo, "Imprimir");
                                    }
                                }

                            }

                        }
                        else
                        {
                            RokuSN SN = new RokuSN();
                            List<String> Seriais = SN.get_Roku_Label(quantidade, modelo);
                            //
                            if (Seriais.Count <= 0)
                            {
                                //para emitir som de alerta
                                Som objSom = new Som();
                                objSom.Falha();
                                //   
                                Clear_ProgressBar();
                                MessageBox.Show("Erro ao gerar seriais", "", MessageBoxButtons.OK, MessageBoxIcon.Error);

                                return;
                            }
                            //
                            if (Seriais.Count > 0)
                            {

                                int de = 1;
                                int ate = Seriais.Count;

                                for (int index = 0; index < ate; index++)
                                {
                                    //VALOR BARRA DE STATUS
                                    createNewRecords(de, ate);

                                    string serial = Seriais[index].ToString();
                                    List<String> Sn = new List<String>();
                                    Sn.Add(serial);

                                    //LABEL codesoft
                                    Print codesoft = new Print();
                                    codesoft.Etiqueta_CodeSoft(Sn, modelo, quantidade, quantidade, etiquetaDupla, impressora, snHorizontal, snVertical, "");

                                    //LOG
                                    Log objLog = new Log();
                                    objLog.Gravar(serial, modelo, "Imprimir");
                                }
                            }

                        }

                        #endregion

                    }
                    //else if (modelo.Equals("ROKU-SMT"))
                    //{
                    //    #region ROKU SMT

                    //    RokuSN SN = new RokuSN();
                    //    List<String> Seriais = SN.get_Roku_Label(quantidade, modelo);
                    //    //
                    //    if (Seriais.Count <= 0)
                    //    {
                    //        //para emitir som de alerta
                    //        Som objSom = new Som();
                    //        objSom.Falha();
                    //        //   
                    //        Clear_ProgressBar();
                    //        MessageBox.Show("Erro ao gerar seriais", "", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    //        return;
                    //    }
                    //    //
                    //    if (Seriais.Count > 0)
                    //    {

                    //        int de = 1;
                    //        int ate = Seriais.Count;

                    //        for (int index = 0; index < ate; index++)
                    //        {
                    //            //VALOR BARRA DE STATUS
                    //            createNewRecords(de, ate);

                    //            string serial = Seriais[index].ToString();
                    //            List<String> Sn = new List<String>();
                    //            Sn.Add(serial);

                    //            //LABEL codesoft
                    //            Print codesoft = new Print();
                    //            codesoft.Etiqueta_CodeSoft(Sn, modelo, quantidade, quantidade, etiquetaDupla, impressora, snHorizontal, snVertical);

                    //            //LOG
                    //            Log objLog = new Log();
                    //            objLog.Gravar(serial, modelo, "Imprimir");
                    //        }
                    //    }


                    //    #endregion
                    //}
                    else if (modelo.Equals("ROKU-GIFIT_BOX"))
                    {
                        #region ROKU GIFT_BOX

                        string mac = txtValor.Text.Trim().ToUpper();

                        if (string.IsNullOrEmpty(mac))
                        {
                            //para emitir som de alerta
                            Som objSom = new Som();
                            objSom.Falha();
                            //   
                            Clear_ProgressBar();
                            MessageBox.Show("Informe número do MAC", "", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            return;
                        }

                        RokuSN Label = new RokuSN();
                        List<RokuSN.Gifit> Seriais = Label.Gifit_Box(mac);
                        //
                        if (Seriais.Count > 0)
                        {

                            int de = 1;
                            int ate = quantidade;//Seriais.Count;

                            for (int index = 0; index < ate; index++)
                            {
                                //VALOR BARRA DE STATUS
                                createNewRecords(de, ate);

                                string serial = Seriais[0].MAC;//Seriais[index].ToString();
                                string wo = Seriais[0].LOT;//Seriais[index].ToString();
                                List<String> Sn = new List<String>();
                                Sn.Add(serial);
                                Sn.Add(wo);

                                //LABEL codesoft
                                Print codesoft = new Print();
                                codesoft.Etiqueta_CodeSoft(Sn, modelo, quantidade, quantidade, etiquetaDupla, impressora, snHorizontal, snVertical, "");

                                //LOG
                                Log objLog = new Log();
                                objLog.Gravar(serial, modelo, "Imprimir");
                            }
                        }
                        else
                        {
                            //para emitir som de alerta
                            Som objSom = new Som();
                            objSom.Falha();
                            //   
                            Clear_ProgressBar();
                            MessageBox.Show("Erro ao gerar seriais", "", MessageBoxButtons.OK, MessageBoxIcon.Error);

                        }


                        #endregion
                    }
                    else if (modelo.Equals("ROKU-PANEL"))//ROKU-BLANK
                    {
                        #region ROKU PANEL

                        if (etiquetaDupla)
                        {
                            if (quantidade % 2 != 0)
                            {
                                //para emitir som de alerta
                                Som objSom = new Som();
                                objSom.Falha();
                                // 
                                Clear_ProgressBar();
                                MessageBox.Show("Quantidade de Impressão somente números pares", "", MessageBoxButtons.OK, MessageBoxIcon.Error);

                                return;
                            }
                            else
                            {
                                RokuSN SN = new RokuSN();
                                List<String> Seriais = SN.get_Roku_Label(quantidade, modelo);
                                //
                                if (Seriais.Count <= 0)
                                {
                                    //para emitir som de alerta
                                    Som objSom = new Som();
                                    objSom.Falha();
                                    // 
                                    Clear_ProgressBar();
                                    MessageBox.Show("Erro ao gerar seriais", "", MessageBoxButtons.OK, MessageBoxIcon.Error);

                                    return;
                                }

                                List<String> Lista_Serial_Duplo = new List<String>();
                                string ladoEsquerdo = string.Empty;
                                string ladoDireito = string.Empty;
                                //
                                for (int index = 0; index < quantidade; index++)
                                {
                                    if (index == 0)
                                    {
                                        ladoEsquerdo = Seriais[index].ToString();
                                    }
                                    else
                                    {
                                        if (index % 2 == 0)//par
                                            ladoEsquerdo = Seriais[index].ToString();
                                        if (index % 2 != 0)//impar
                                            ladoDireito = Seriais[index].ToString();
                                    }

                                    if ((ladoEsquerdo != string.Empty) && (ladoDireito != string.Empty))
                                    {
                                        string sn1_sn2 = ladoEsquerdo + ";" + ladoDireito;
                                        Lista_Serial_Duplo.Add(sn1_sn2);
                                        //
                                        ladoEsquerdo = string.Empty;
                                        ladoDireito = string.Empty;
                                    }

                                }
                                //
                                if (Lista_Serial_Duplo.Count > 0)
                                {
                                    int de = 1;
                                    int ate = Lista_Serial_Duplo.Count;
                                    //
                                    for (int index = 0; index < ate; index++)
                                    {
                                        //VALOR BARRA DE STATUS
                                        createNewRecords(de, ate);

                                        string serial = Lista_Serial_Duplo[index].ToString();
                                        List<String> sn = new List<String>();
                                        sn.Add(serial);

                                        //LABEL codesoft
                                        Print codesoft = new Print();
                                        codesoft.Etiqueta_CodeSoft(sn, modelo, 1, quantidade, etiquetaDupla, impressora, snHorizontal, snVertical, "");

                                        //LOG                                    
                                        Log objLog = new Log();
                                        objLog.Gravar(serial, modelo, "Imprimir");
                                    }
                                }

                            }

                        }

                        #endregion
                    }
                    else
                    {
                        if (modelo.Equals("TG3442-SMT"))
                        {
                            //3 últimos digitos da wo
                            wo_Last3digits = txtValor.Text.ToUpper().Trim();
                            if (!string.IsNullOrEmpty(wo_Last3digits))
                            {
                                string aux = "";
                                for (int index = 3; index >= 1; index--)
                                {
                                    aux += wo_Last3digits[wo_Last3digits.Length - index];
                                }
                                //
                                wo_Last3digits = aux;

                            }
                            //4 últimos digitos da wo
                            wo_Last4digits = txtValor.Text.ToUpper().Trim();
                            if (!string.IsNullOrEmpty(wo_Last4digits))
                            {
                                string aux = "";
                                for (int index = 4; index >= 1; index--)
                                {
                                    aux += wo_Last4digits[wo_Last4digits.Length - index];
                                }
                                //
                                wo_Last4digits = aux;

                            }

                        }

                        if (etiquetaDupla)
                        {
                            if (quantidade % 2 != 0)
                            {
                                //para emitir som de alerta
                                Som objSom = new Som();
                                objSom.Falha();
                                // 
                                Clear_ProgressBar();
                                MessageBox.Show("Quantidade de Impressão somente números pares", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                            else
                            {
                                ARRIS_SN SN = new ARRIS_SN();
                                List<String> Seriais = SN.get_Arris_SN(quantidade, modelo, wo_Last3digits);
                                //
                                if (Seriais.Count <= 0)
                                {
                                    //para emitir som de alerta
                                    Som objSom = new Som();
                                    objSom.Falha();
                                    // 
                                    Clear_ProgressBar();
                                    MessageBox.Show("Erro ao gerar seriais", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return;
                                }

                                List<String> Lista_Serial_Duplo = new List<String>();
                                string ladoEsquerdo = string.Empty;
                                string ladoDireito = string.Empty;
                                //
                                for (int index = 0; index < quantidade; index++)
                                {
                                    if (index == 0)
                                    {
                                        ladoEsquerdo = Seriais[index].ToString();
                                    }
                                    else
                                    {
                                        if (index % 2 == 0)//par
                                            ladoEsquerdo = Seriais[index].ToString();
                                        if (index % 2 != 0)//impar
                                            ladoDireito = Seriais[index].ToString();
                                    }

                                    if ((ladoEsquerdo != string.Empty) && (ladoDireito != string.Empty))
                                    {
                                        string sn1_sn2 = ladoEsquerdo + ";" + ladoDireito;
                                        Lista_Serial_Duplo.Add(sn1_sn2);
                                        //
                                        ladoEsquerdo = string.Empty;
                                        ladoDireito = string.Empty;
                                    }

                                }
                                //
                                if (Lista_Serial_Duplo.Count > 0)
                                {
                                    int de = 1;
                                    int ate = Lista_Serial_Duplo.Count;
                                    //
                                    for (int index = 0; index < ate; index++)
                                    {
                                        //VALOR BARRA DE STATUS
                                        createNewRecords(de, ate);

                                        string serial = Lista_Serial_Duplo[index].ToString();
                                        List<String> sn = new List<String>();
                                        sn.Add(serial);

                                        //LABEL codesoft
                                        Print codesoft = new Print();
                                        codesoft.Etiqueta_CodeSoft(sn, modelo, 1, quantidade, etiquetaDupla, impressora, snHorizontal, snVertical, wo_Last4digits);

                                        //LOG                                    
                                        Log objLog = new Log();
                                        objLog.Gravar(serial, modelo, "Imprimir");
                                    }
                                }

                            }

                        }
                        else
                        {
                            ARRIS_SN SN = new ARRIS_SN();
                            List<String> Seriais = SN.get_Arris_SN(quantidade, modelo, wo_Last3digits);
                            //
                            if (Seriais.Count <= 0)
                            {
                                //para emitir som de alerta
                                Som objSom = new Som();
                                objSom.Falha();
                                //   
                                Clear_ProgressBar();
                                MessageBox.Show("Erro ao gerar seriais", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                            //
                            if (Seriais.Count > 0)
                            {
                                int de = 1;
                                int ate = Seriais.Count;
                                //
                                for (int index = 0; index < ate; index++)
                                {
                                    //VALOR BARRA DE STATUS
                                    createNewRecords(de, ate);

                                    string serial = Seriais[index].ToString();
                                    List<String> Sn = new List<String>();
                                    Sn.Add(serial);

                                    //LABEL codesoft
                                    Print codesoft = new Print();
                                    codesoft.Etiqueta_CodeSoft(Sn, modelo, quantidade, quantidade, etiquetaDupla, impressora, snHorizontal, snVertical, wo_Last4digits);

                                    //LOG
                                    Log objLog = new Log();
                                    objLog.Gravar(serial, modelo, "Imprimir");
                                }
                            }

                        }
                    }
                }
                else
                {
                    ARRIS_SN SN = new ARRIS_SN();
                    List<String> Seriais = SN.get_Arris_SN(quantidade, modelo, string.Empty);
                    //
                    if (Seriais.Count <= 0)
                    {
                        //para emitir som de alerta
                        Som objSom = new Som();
                        objSom.Falha();
                        //  
                        Clear_ProgressBar();
                        MessageBox.Show("Erro ao gerar seriais", "", MessageBoxButtons.OK, MessageBoxIcon.Error);

                        return;
                    }
                    //
                    int index_total_Impressao = 0;
                    //
                    if (modelo.Equals("E965") || modelo.Equals("DSI"))
                    {
                        if (Seriais.Count > 0)
                        {
                            //Thread.CurrentThread.GetApartmentState();
                            int de = 1;
                            int ate = Seriais.Count;
                            //
                            for (int linha = 0; linha < ate; linha++)
                            {
                                //VALOR BARRA DE STATUS
                                createNewRecords(de, ate);
                                //
                                List<String> Sn = new List<String>();
                                string serial_ = Seriais[linha].ToString();
                                Sn.Add(serial_);

                                //GEERAR E IMPRIMIR ETIQUETA
                                Print codesoft = new Print();
                                codesoft.Etiqueta_CodeSoft(Sn, modelo, quantidade, quantidade, etiquetaDupla, impressora, string.Empty, string.Empty, "");

                                //LOG
                                string serial = Seriais[linha].ToString();
                                Log objLog = new Log();
                                objLog.Gravar(serial, modelo, "Imprimir");

                            }

                        }

                    }
                    else//outros modelos
                    {
                        foreach (String serial in Seriais)
                        {
                            index_total_Impressao++;

                            //LABEL ZPL 
                            Print zpl = new Print();
                            zpl.Imprimir(serial, modelo, quantidade, index_total_Impressao, etiquetaDupla, UsarImpressoraPadrao, impressora);
                            //
                            Log objLog = new Log();
                            objLog.Gravar(serial, modelo, "Imprimir");
                        }
                    }

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
        }
        //
        private void txtQty_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void cboModelo_SelectedValueChanged(object sender, EventArgs e)
        {
            string modelo = cboModelo.SelectedItem.ToString();
            txtValor.Clear();
            //
            if ((modelo.Equals("TG1692BR-SMT")) || (modelo.Equals("TG1692A-SMT")))
            {
                chkEtiquetaDupla.Enabled = true;
                chkEtiquetaDupla.Checked = false;
                //                
                lbSnVertical.Visible = false;
                txtValor.Visible = false;

            }
            else if (modelo.Equals("TG3442-SMT"))
            {
                //chkEtiquetaDupla.Enabled = true;
                chkEtiquetaDupla.Checked = false;
                chkEtiquetaDupla.Enabled = true;
                //  
                lbSnVertical.Text = "WO:";
                lbSnVertical.Visible = true;
                txtValor.Visible = true;
            }
            else if (modelo.Equals("ROKU-MAC"))
            {
                chkEtiquetaDupla.Enabled = true;
                //chkEtiquetaDupla.Checked = false;
                //chkEtiquetaDupla.Enabled = false;
                //  
                lbSnVertical.Text = "WO:";
                lbSnVertical.Visible = true;
                txtValor.Visible = true;
            }
            else if (modelo.Equals("ROKU-GIFIT_BOX"))
            {
                lbSnVertical.Text = "MAC:";
                lbSnVertical.Visible = true;
                txtValor.Visible = true;
            }
            else if (modelo.Equals("ROKU-PANEL"))
            {
                chkEtiquetaDupla.Enabled = false;
                chkEtiquetaDupla.Checked = true;
                //                
                lbSnVertical.Visible = false;
                txtValor.Visible = false;
            }
            else if (modelo.Equals("ROKU-SMT"))
            {
                chkEtiquetaDupla.Enabled = true;//false;
                //chkEtiquetaDupla.Checked = false;
                //                
                lbSnVertical.Text = "WO:";
                lbSnVertical.Visible = true;
                txtValor.Visible = true;
            }
            else
            {
                chkEtiquetaDupla.Enabled = false;
                chkEtiquetaDupla.Checked = false;
                //                
                lbSnVertical.Visible = false;
                txtValor.Visible = false;

            }
            //
            txtValor.Focus();
        }

    }
}
