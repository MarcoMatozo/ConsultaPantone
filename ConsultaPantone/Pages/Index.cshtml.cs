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
                            CorCo = SafeDouble(reader["corco"]),
                            CorCoAlt = SafeDouble(reader["corcoalt"]),
                            CorPa = SafeDouble(reader["corpa"]),
                            CorCv = SafeDouble(reader["corcv"]), // Verifique se no banco é corcv ou corpv
                            CorPes = SafeDouble(reader["corpes"]),
                            CorPoli = SafeDouble(reader["corpoli"]),
                            Red = SafeInt(reader["red"]),
                            Green = SafeInt(reader["green"]),
                            Blue = SafeInt(reader["blue"]),
                            // Carrega os campos de verificação
                            Verificada = reader["verificada"] != DBNull.Value && Convert.ToBoolean(reader["verificada"]),
                            Verificadapes = reader["verificadapes"] != DBNull.Value && Convert.ToBoolean(reader["verificadapes"])
                            // Adicione outros campos se desejar exibir mais
                        });
                    }
                }
            }
            double SafeDouble(object value) => value == DBNull.Value ? 0 : Convert.ToDouble(value);
            int SafeInt(object value) => value == DBNull.Value ? 0 : Convert.ToInt32(value);
        }

        // MÉTODO PARA CADASTRAR OU EDITAR
        public IActionResult OnPostSalvar(PantoneItem item)
        {
            string connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=P:\Laboratorio\01 - Lab-PUB\NovoPantone\PantoneConsulta_be.accdb;";

            using (OleDbConnection conn = new OleDbConnection(connString))
            {
                conn.Open();

                // Verifica se o registro já existe para decidir entre INSERT ou UPDATE
                string checkSql = "SELECT COUNT(*) FROM TABpan WHERE pantonetpx = ?";
                OleDbCommand checkCmd = new OleDbCommand(checkSql, conn);
                checkCmd.Parameters.AddWithValue("?", item.PantoneTpx ?? "");
                int existe = (int)checkCmd.ExecuteScalar();

                string sql;
                if (existe > 0)
                {
                    sql = @"UPDATE TABpan SET nomepantone=?, pagina2=?, coluna2=?, linha2=?, 
                    corco=?, corcoalt=?, corpa=?, corcv=?, corpes=?, corpoli=?, 
                    red=?, green=?, blue=?, verificada=?, verificadapes=? 
                    WHERE pantonetpx=?";
                }
                else
                {
                    sql = @"INSERT INTO TABpan (nomepantone, pagina2, coluna2, linha2, 
                    corco, corcoalt, corpa, corcv, corpes, corpoli, 
                    red, green, blue, verificada, verificadapes, pantonetpx) 
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
                }

                OleDbCommand cmd = new OleDbCommand(sql, conn);
                // A ORDEM DOS PARÂMETROS DEVE SER IDÊNTICA AO SQL ACIMA
                cmd.Parameters.AddWithValue("?", item.NomePantone ?? "");
                cmd.Parameters.AddWithValue("?", item.Pagina2 ?? "");
                cmd.Parameters.AddWithValue("?", item.Coluna2 ?? "");
                cmd.Parameters.AddWithValue("?", item.Linha2 ?? "");
                cmd.Parameters.AddWithValue("?", item.CorCo);
                cmd.Parameters.AddWithValue("?", item.CorCoAlt);
                cmd.Parameters.AddWithValue("?", item.CorPa);
                cmd.Parameters.AddWithValue("?", item.CorCv);
                cmd.Parameters.AddWithValue("?", item.CorPes);
                cmd.Parameters.AddWithValue("?", item.CorPoli);
                cmd.Parameters.AddWithValue("?", item.Red);
                cmd.Parameters.AddWithValue("?", item.Green);
                cmd.Parameters.AddWithValue("?", item.Blue);
                cmd.Parameters.AddWithValue("?", item.Verificada);
                cmd.Parameters.AddWithValue("?", item.Verificadapes);
                cmd.Parameters.AddWithValue("?", item.PantoneTpx ?? "");

                cmd.ExecuteNonQuery();
            }
            return RedirectToPage();
        }

        public IActionResult OnPostExcluir(string id)
        {
            string connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=P:\Laboratorio\01 - Lab-PUB\NovoPantone\PantoneConsulta_be.accdb;";
            using (OleDbConnection conn = new OleDbConnection(connString))
            {
                string sql = "DELETE FROM TABpan WHERE pantonetpx = ?";
                OleDbCommand cmd = new OleDbCommand(sql, conn);
                cmd.Parameters.AddWithValue("?", id);
                conn.Open();
                cmd.ExecuteNonQuery();
            }
            return RedirectToPage();
        }

    }
}