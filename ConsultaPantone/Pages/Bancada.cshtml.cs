using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Data.OleDb;
using System.Text.Json;
using System.Globalization;

namespace ConsultaPantone.Pages
{
    public class BancadaModel : PageModel
    {
        // AJUSTE O CAMINHO DO SEU BANCO DE DADOS ABAIXO
        private readonly string _connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=P:\Laboratorio\01 - Lab-PUB\NovoPantone\PantoneConsulta_be.accdb;";

        [BindProperty]
        public string DadosJson { get; set; }

        [BindProperty]
        public int Codigocliente { get; set; }

        [BindProperty]
        public string Nomecliente { get; set; }

        public void OnGet() { }

        // Busca o nome do cliente para exibir na tela
        public JsonResult OnGetBuscarCliente(int cod)
        {
            string nomeEncontrado = "";
            try
            {
                using (OleDbConnection conn = new OleDbConnection(_connectionString))
                {
                    string sql = "SELECT nomecliente FROM TABcliente WHERE codcliente = ?";
                    using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                    {
                        cmd.Parameters.Add("?", OleDbType.Integer).Value = cod;
                        conn.Open();
                        var result = cmd.ExecuteScalar();
                        if (result != null) nomeEncontrado = result.ToString();
                    }
                }
                return new JsonResult(new { sucesso = !string.IsNullOrEmpty(nomeEncontrado), nome = nomeEncontrado });
            }
            catch (Exception ex)
            {
                return new JsonResult(new { sucesso = false, erro = ex.Message });
            }
        }

        public IActionResult OnPostSalvarEnvio()
        {
            if (string.IsNullOrEmpty(DadosJson) || Codigocliente <= 0) return Page();

            try
            {
                var itens = JsonSerializer.Deserialize<List<ItemEnvio>>(DadosJson);

                using (OleDbConnection conn = new OleDbConnection(_connectionString))
                {
                    conn.Open();
                    foreach (var item in itens)
                    {
                        // SQL incluindo o campo nomecliente
                        string sql = @"INSERT INTO TABenvios 
                                     (codigocliente, nomecliente, PantoneTpx, DataHora, CorCo, CorPes, CorPA, CorPoli, CorCoAlt) 
                                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)";

                        using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                        {
                            // A ORDEM DOS PARÂMETROS DEVE SER IDÊNTICA AO SQL ACIMA
                            cmd.Parameters.Add("?", OleDbType.Integer).Value = Codigocliente;
                            cmd.Parameters.Add("?", OleDbType.VarWChar).Value = (object)Nomecliente ?? DBNull.Value;
                            cmd.Parameters.Add("?", OleDbType.VarWChar).Value = (object)item.PantoneTpx ?? DBNull.Value;
                            cmd.Parameters.Add("?", OleDbType.Date).Value = DateTime.Now;

                            // Valores numéricos das fórmulas
                            cmd.Parameters.Add("?", OleDbType.Double).Value = TentarConverter(item.co);
                            cmd.Parameters.Add("?", OleDbType.Double).Value = TentarConverter(item.pes);
                            cmd.Parameters.Add("?", OleDbType.Double).Value = TentarConverter(item.pa);
                            cmd.Parameters.Add("?", OleDbType.Double).Value = TentarConverter(item.poli);
                            cmd.Parameters.Add("?", OleDbType.Double).Value = TentarConverter(item.alt);

                            cmd.ExecuteNonQuery();
                        }
                    }
                }
                return RedirectToPage("Index");
            }
            catch (Exception ex)
            {
                throw new Exception("Erro ao salvar histórico: " + ex.Message);
            }
        }

        private object TentarConverter(string valor)
        {
            if (string.IsNullOrWhiteSpace(valor)) return DBNull.Value;
            string valorLimpo = valor.Replace(",", ".");
            if (double.TryParse(valorLimpo, NumberStyles.Any, CultureInfo.InvariantCulture, out double resultado))
            {
                return resultado;
            }
            return DBNull.Value;
        }

        public class ItemEnvio
        {
            public string PantoneTpx { get; set; }
            public string co { get; set; }
            public string pes { get; set; }
            public string pa { get; set; }
            public string poli { get; set; }
            public string alt { get; set; }
        }
    }
}