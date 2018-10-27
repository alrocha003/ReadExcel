using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;

namespace ReadExcel
{
    public class Read
    {
        public string Path { get; private set; }
        public DataSet Ds { get; private set; }

        public Read(string _path)
        {
            this.Path = _path;
            ReadArq();
        }

        public void ReadArq()
        {
            OleDbConnection conexao = new OleDbConnection($@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = 
            {Path};Extended Properties ='Excel 12.0 Xml;HDR=YES';");
            OleDbDataAdapter adapter = new OleDbDataAdapter("select * from[tb_produto$]", conexao);
            this.Ds = new DataSet();

            try
            {
                conexao.Open();
                adapter.Fill(Ds);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao acessar os dados: {ex.Message}");
            }
            finally
            {
                conexao.Close();
            }
        }

        public void CreateJson()
        {
            int i = 0;
            int rowId = 0;
            StringBuilder query = new StringBuilder();

            if (this.Ds != null)
            {
                foreach (DataColumn item in this.Ds.Tables[0].Columns)
                {
                    query.AppendLine($@"{(i == 0 ? "{" : "")}{'"'}{item.ColumnName}{'"'}: 
                                    {(Ds.Tables[0].Rows[rowId][i].GetType() == typeof(string) ? ('"' + Ds.Tables[0].Rows[rowId][i].ToString() + '"')
                                    : Ds.Tables[0].Rows[rowId][i])}{(i == 1 ? "" : ",")}");
                    i++;
                    rowId++;
                }
                query.AppendLine("}");
                Console.WriteLine(query);
                Console.ReadLine();
            }
        }

        public void ReadRows()
        {

            if (this.Ds != null)
            {
                foreach (DataRow linha in this.Ds.Tables[0].Rows)
                {
                    Console.WriteLine($@"Cód. Produto: {linha["CD_PRODUTO"].ToString()} – Descricao: {linha["DESCRICAO"].ToString()} 
                                        – Preço: {linha["PRECO"].ToString()} - Quantidade: {linha["QTDE"].ToString()}");
                }
                Console.ReadLine();
            }
        }
    }
}
