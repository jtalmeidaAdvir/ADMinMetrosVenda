using System;
using System.Diagnostics;
using System.Globalization;
using System.Windows; // WinForms MessageBox (estável no Evolution)
using Primavera.Extensibility.BusinessEntities.ExtensibilityService.EventArgs;
using Primavera.Extensibility.Sales.Editors;

namespace ADMinMetrosVenda
{
    public class PriClass1 : EditorVendas
    {
        // Ajusta estes flags conforme pretendes.
        private const bool FORCAR_SEMPRE = true; // true = força sempre; false = só quando vazio/1
        private const double EPS = 1e-6;         // tolerância para comparação de doubles

        // Ativa logs definindo o símbolo de compilação: DEBUG_ADM
        [Conditional("DEBUG_ADM")]
        private static void Dbg(string msg)
        {
            try { MessageBox.Show(msg, "DEBUG ADMinMetrosVenda"); } catch { /* ignore */ }
        }

        public override void ArtigoIdentificado(string Artigo, int NumLinha, ref bool Cancel, ExtensibilityEventArgs e)
        {
            //Dbg($"Entrou em ArtigoIdentificado | Artigo={Artigo} | NumLinha={NumLinha}");

            // Validar inputs mais baratos primeiro
            if (string.IsNullOrWhiteSpace(Artigo))
                return;

            if (NumLinha < 0)
                return;

            // Obter MinMetros do CDU
            if (!TryGetMinMetros(Artigo, out var minMetros))
            {
                //Dbg("MinMetros não disponível/<=0.");
                return;
            }

            if (minMetros <= 0.0)
                return;

            // Obter coleção de linhas do documento
            dynamic linhas = this.DocumentoVenda?.Linhas;
            if (linhas == null)
                return;

            // Obter a linha de forma segura
            if (!SafeGetLinha(linhas, NumLinha, out dynamic linha))
            {
                //Dbg("Não foi possível obter a linha.");
                return;
            }

            // Ler quantidade atual
            double qtdAntes = 0.0;
            try { qtdAntes = (double)linha.Quantidade; } catch { /* deixa a 0 */ }

            // Decisão de alteração
            bool deveForcar =
                FORCAR_SEMPRE ||
                qtdAntes <= 0.0 + EPS ||
                Math.Abs(qtdAntes - 1.0) < EPS;

            if (!deveForcar)
            {
                //Dbg($"Mantida quantidade existente ({qtdAntes}).");
                return;
            }

            // Opcional: arredondar se precisares de limitar casas decimais (ex.: 3)
            // minMetros = Math.Round(minMetros, 3, MidpointRounding.AwayFromZero);

            try
            {
                linha.Quantidade = minMetros;
                //Dbg($"Quantidade alterada: {qtdAntes} → {minMetros}");
            }
            catch (Exception ex)
            {
                Dbg("Falha ao atribuir Quantidade: " + ex.Message);
            }
        }

        /// <summary>
        /// Obtém o valor do CDU_MinMetrosSugestaoVenda como double (InvariantCulture) de forma segura.
        /// </summary>
        private bool TryGetMinMetros(string artigo, out double minMetros)
        {
            minMetros = 0.0;

            // Escapar simples para evitar quebra de SQL
            string artigoCod = artigo.Replace("'", "''");
            string sql = $"SELECT CDU_MinMetrosSugestaoVenda AS MinMetros FROM Artigo WHERE Artigo = '{artigoCod}'";

            dynamic rs = null;
            try { rs = BSO?.Consulta(sql); } catch (Exception ex) { Dbg("Consulta falhou: " + ex.Message); return false; }

            if (rs == null)
                return false;

            try
            {
                if (rs.Vazia())
                    return false;

                object raw = rs.Valor("MinMetros");
                if (raw == null || raw is DBNull)
                    return false;

                // Tentar parse direto (mais rápido se vier como string)
                if (raw is string s)
                {
                    if (double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var v1))
                    {
                        minMetros = v1;
                        return true;
                    }
                    // Tentar cultura atual como fallback (caso alguém grave com vírgula)
                    if (double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out var v2))
                    {
                        minMetros = v2;
                        return true;
                    }
                }

                // Caso venha numérico (decimal/double/float/etc.)
                minMetros = Convert.ToDouble(raw, CultureInfo.InvariantCulture);
                return true;
            }
            catch (Exception ex)
            {
                Dbg("Erro a ler/convertar MinMetros: " + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Obtém a linha com GetEdita de forma segura (evita lançar exceções).
        /// </summary>
        private bool SafeGetLinha(dynamic linhas, int numLinha, out dynamic linha)
        {
            linha = null;
            try
            {
                linha = linhas.GetEdita(numLinha);
                return linha != null;
            }
            catch (Exception ex)
            {
                Dbg("GetEdita falhou: " + ex.Message);
                return false;
            }
        }
    }
}
