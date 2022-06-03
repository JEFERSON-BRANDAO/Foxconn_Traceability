using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Collections;
using System.Data.SqlClient;
using System.Data.OracleClient;

namespace Classes
{
    public class OracleConnect
    {
        #region Attributes

        private bool _isvalid;
        private string _message;
        private string _stringConnection;
        private OracleConnection _connection;
        private DataTable _tabela;
        private IList _parametros = new ArrayList();
        private OracleTransaction _transaction;
        private OracleCommand _command;

        #endregion
        //
        #region Properties

        public string StringConnection
        {
            get { return _stringConnection; }
            set { _stringConnection = value; }
        }

        public OracleConnection Connection
        {
            get { return _connection; }
            set { _connection = value; }
        }

        public DataTable Tabela
        {
            get { return _tabela; }
            set { _tabela = value; }
        }

        public IList Parametros
        {
            get { return _parametros; }
            set { _parametros = value; }
        }

        public OracleTransaction Transaction
        {
            get { return _transaction; }
            set { _transaction = value; }
        }

        public bool Isvalid
        {
            get { return _isvalid; }
            set { _isvalid = value; }
        }

        public string Message
        {
            get { return _message; }
            set { _message = value; }
        }

        #endregion
        //
        #region Methods

        public bool Conectar()
        {
            try
            {
                Criptografia objCript = new Criptografia();

                //valores guardados em Settings.settings
                string DATABASE = objCript.Descriptografar(Foxconn_Traceability.Properties.Settings.Default.DataSource);
                string USUARIO = objCript.Descriptografar(Foxconn_Traceability.Properties.Settings.Default.Usuario);
                string SENHA = objCript.Descriptografar(Foxconn_Traceability.Properties.Settings.Default.Senha);

                string _connectionString = "Data Source=" + DATABASE + "; Persist Security Info=True;User ID=" + USUARIO + ";Password=" + SENHA + ";Unicode=True";
                //string _connectionString = "Provider=OraOLEDB.Oracle;Data Source=" + DATABASE + ";User Id=" + USUARIO + ";Password=" + SENHA + ";";

                _connection = new OracleConnection(_connectionString);

                _connection.Open();
                return true;
            }
            catch (Exception erro)
            {
                Message = erro.Message;
                return false;
            }
        }

        public void Desconectar()
        {
            try
            {
                _connection.Close();
            }
            catch (Exception) { }
        }

        public void AdicionarParametro(string nome, object valor, OracleType tipo)//SqlDbType
        {
            OracleParameter parametro = new OracleParameter(nome, tipo);
            parametro.Direction = ParameterDirection.Input;
            parametro.Value = valor;

            _parametros.Add(parametro);
        }

        public void AdicionarParametroSaida(string nome, OracleType tipo)//SqlDbType
        {
            OracleParameter parametro = new OracleParameter(nome, tipo);
            parametro.Direction = ParameterDirection.Output;//ParameterDirection.ReturnValue;
            parametro.Size = 50;

            _parametros.Add(parametro);
        }

        public void SetarSQL(string SQL)
        {
            _command = new OracleCommand();
            _command.CommandType = CommandType.Text;
            _command.CommandText = SQL;
            _command.Connection = _connection;
        }

        public void SetarSP(string nomeSP)
        {
            _command = new OracleCommand();
            _command.CommandType = CommandType.StoredProcedure;
            _command.CommandText = nomeSP;
            _command.Connection = _connection;
        }

        public bool Executar()
        {

            try
            {
                //_command.Parameters.Clear();

                foreach (OracleParameter parametro in _parametros)
                {
                    _command.Parameters.Add(parametro);

                    //_command.Parameters.AddWithValue(parametro.ToString(), "");
                }

                //_parametros = new ArrayList();

                OracleDataAdapter dataAdapter = new OracleDataAdapter(_command);
                Tabela = new DataTable();
                dataAdapter.Fill(Tabela);

                _isvalid = true;
                _message = "";

                return true;
            }
            catch (Exception erro)
            {
                _isvalid = false;
                _message = erro.Message;

                return false;
            }
        }

        public bool ExecutarProcedure()
        {
            try
            {
                foreach (OracleParameter parametro in _parametros)
                {
                    _command.Parameters.Add(parametro);
                }

                /*ESSE METODO NAO FUNCIONA PARAMETRO DE RETORNO POIS SO VEM VALOR. NAO TEM COLUNA PARA SER INSERIDO NO DataTable*/
                //OracleDataAdapter dataAdapter = new OracleDataAdapter(_command);
                //Tabela = new DataTable();
                //dataAdapter.Fill(Tabela);

                _command.ExecuteReader();
                string retorno = _command.Parameters["VAR_RES"].Value.ToString();
                Tabela = new DataTable();
                //
                #region Preenche DataTable

                DataColumn column = new DataColumn();
                column.DataType = System.Type.GetType("System.String");
                column.ColumnName = "SERIAL";
                Tabela.Columns.Add(column);
                //
                DataRow row = Tabela.NewRow();
                row["SERIAL"] = retorno;
                Tabela.Rows.Add(row);

                #endregion

                //
                _isvalid = true;
                _message = "";
                //
                return true;
            }
            catch (Exception erro)
            {
                _isvalid = false;
                _message = erro.Message;

                return false;
            }
        }


        #endregion
    }
}
