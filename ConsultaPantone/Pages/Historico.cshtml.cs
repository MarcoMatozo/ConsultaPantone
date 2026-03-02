using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Data.OleDb;
using System.Globalization;

namespace ConsultaPantone.Pages
{
    public class HistoricoModel : PageModel
    {
        // Certifique-se de que o caminho aponta para o seu banco de dados
        private readonly string _connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=P:\Laboratorio\01 - Lab-PUB\NovoPantone\PantoneConsulta_be.accdb;";

        public List<EnvioRegisto> Envios { get; set; } = new List<EnvioRegisto>();

        [BindProperty(SupportsGet = true)]
        public string Filtro { get; set; }

        public void OnGet()
        {
            using (OleDbConnection conn = new OleDbConnection(_connectionString))
            {
                // SQL com busca em 3 campos: Nome, Código Cliente ou Pantone
                string sql = "SELECT * FROM TABenvios";

                if (!string.IsNullOrEmpty(Filtro))
                {
                    sql += " WHERE nomecliente LIKE ? OR codigocliente LIKE ? OR PantoneTpx LIKE ?";
                }

                sql += " ORDER BY DataHora DESC";

                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    if (!string.IsNullOrEmpty(Filtro))
                    {
                        // O OleDb exige um parâmetro para cada '?' na ordem exata
                        string busca = "%" + Filtro + "%";
                        cmd.Parameters.Add("?", OleDbType.VarWChar).Value = busca;
                        cmd.Parameters.Add("?", OleDbType.VarWChar).Value = busca;
                        cmd.Parameters.Add("?", OleDbType.VarWChar).Value = busca;
                    }

                    try
                    {
                        conn.Open();
                        using (OleDbDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                Envios.Add(new EnvioRegisto
                                {
                                    CodCliente = reader["codigocliente"]?.ToString(),
                                    NomeCliente = reader["nomecliente"]?.ToString(),
                                    Pantone = reader["PantoneTpx"]?.ToString(),
                                    Data = reader["DataHora"] != DBNull.Value ? Convert.ToDateTime(reader["DataHora"]) : DateTime.MinValue,
                                    Co = reader["CorCo"]?.ToString(),
                                    Pes = reader["CorPes"]?.ToString(),
                                    Pa = reader["CorPA"]?.ToString(),
                                    Poli = reader["CorPoli"]?.ToString(),
                                    Alt = reader["CorCoAlt"]?.ToString()
                                });
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log de erro se necessário
                        TempData["Erro"] = "Erro ao carregar histórico: " + ex.Message;
                    }
                }
            }
        }

        public class EnvioRegisto
        {
            public string CodCliente { get; set; }
            public string NomeCliente { get; set; }
            public string Pantone { get; set; }
            public DateTime Data { get; set; }
            public string Co { get; set; }
            public string Pes { get; set; }
            public string Pa { get; set; }
            public string Poli { get; set; }
            public string Alt { get; set; }
        }
    }
}