using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Data.OleDb;
using ConsultaPantone.Models;

namespace ConsultaPantone.Pages
{
    public class IndexModel : PageModel
    {
        // Esta lista guardará os resultados da busca
        public List<PantoneItem> Resultados { get; set; } = new();

        [BindProperty(SupportsGet = true)]
        public string? TermoBusca { get; set; }

        public void OnGet()
        {
            if (string.IsNullOrEmpty(TermoBusca)) return;

            // Caminho do seu banco de dados na rede
            string connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=P:\Laboratorio\01 - Lab-PUB\NovoPantone\PantoneConsulta_be.accdb;Persist Security Info=False;";

            using (OleDbConnection conn = new OleDbConnection(connString))
            {
                // Comando SQL para buscar pelo nome ou código TPX
                string sql = "SELECT * FROM TABpan WHERE pantonetpx LIKE ? OR nomepantone LIKE ?";
                OleDbCommand cmd = new OleDbCommand(sql, conn);

                // O '%' serve para buscar partes do texto (ex: buscar "Azul" traz "Azul Royal")
                cmd.Parameters.AddWithValue("?", "%" + TermoBusca + "%");
                cmd.Parameters.AddWithValue("?", "%" + TermoBusca + "%");

                conn.Open();
                using (OleDbDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        Resultados.Add(new PantoneItem
                        {
                            PantoneTpx = reader["pantonetpx"]?.ToString(),
                            NomePantone = reader["nomepantone"]?.ToString(),
                            Pagina2 = reader["pagina2"]?.ToString(),
                            Coluna2 = reader["coluna2"]?.ToString(),
                            Linha2 = reader["linha2"]?.ToString(),

                            // Usamos uma função auxiliar para evitar o erro de DBNull
                            CorCo = SafeDouble(reader["corco"]),
                            CorCoAlt = SafeDouble(reader["corcoalt"]),
                            CorCv = SafeDouble(reader["corcv"]),
                            CorPa = SafeDouble(reader["corpa"]),
                            CorPes = SafeDouble(reader["corpes"]),
                            CorPoli = SafeDouble(reader["corpoli"]),

                            Red = SafeInt(reader["red"]),
                            Green = SafeInt(reader["green"]),
                            Blue = SafeInt(reader["blue"])
                            // Adicione outros campos se desejar exibir mais
                        });
                    }
                }
            }
            double SafeDouble(object value) => value == DBNull.Value ? 0 : Convert.ToDouble(value);
            int SafeInt(object value) => value == DBNull.Value ? 0 : Convert.ToInt32(value);
        }
    }
}