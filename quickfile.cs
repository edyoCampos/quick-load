using GreenDoc.Data;
using GreenDoc.Infrastructure;
using GreenDoc.Models;
using GreenDoc.Services;
using GreenDoc.Services.Helpers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;

namespace GreenDoc.Web.Api.Controllers
{
    [MyBasicAuthenticationFilter]
    public class ItemsController : ApiController
    {
        public IUnitOfWork _unitOfWork;
        public GreenDocContext context;
        public ContaService _contaService;
        public UsuarioService _usuarioService;
        public DocumentoService _documentoService;
        public ProjetoService _projetoService;
        public BPMService _bpmService;
        public ServidorService _servidorService;
        public ArquivoService _arquivoService;
        public PastaService _pastaService;
        public ItensService _itensService;
        public ComentarioService _comentarioService;
        public TriggerService _triggerService;
        public ProcessoService _processoService;
        public ScriptsService _scriptsService;

        public ItemsController()
        {
            context = new GreenDocContext();
            _unitOfWork = new UnitOfWork(context);

            _contaService = new ContaService(_unitOfWork);
            _usuarioService = new UsuarioService(_unitOfWork);
            _triggerService = new TriggerService(_unitOfWork);
            _documentoService = new DocumentoService(_unitOfWork, _triggerService);
            _projetoService = new ProjetoService(_unitOfWork);
            _bpmService = new BPMService(_unitOfWork, _triggerService);
            _servidorService = new ServidorService(_unitOfWork);
            _arquivoService = new ArquivoService(_unitOfWork, _triggerService);
            _pastaService = new PastaService(_unitOfWork);
            _itensService = new ItensService(_unitOfWork);
            _comentarioService = new ComentarioService(_unitOfWork, _triggerService);
            _processoService = new ProcessoService(_unitOfWork);
            _scriptsService = new ScriptsService(_unitOfWork);
        }

        public ItemsController(IUnitOfWork unitOfWork)
        {
            _unitOfWork = unitOfWork;

            _contaService = new ContaService(_unitOfWork);
            _usuarioService = new UsuarioService(_unitOfWork);
            _triggerService = new TriggerService(_unitOfWork);
            _documentoService = new DocumentoService(_unitOfWork, _triggerService);
            _projetoService = new ProjetoService(_unitOfWork);
            _bpmService = new BPMService(_unitOfWork, _triggerService);
            _servidorService = new ServidorService(_unitOfWork);
            _arquivoService = new ArquivoService(_unitOfWork, _triggerService);
            _pastaService = new PastaService(_unitOfWork);
            _itensService = new ItensService(_unitOfWork);
            _comentarioService = new ComentarioService(_unitOfWork, _triggerService);
            _processoService = new ProcessoService(_unitOfWork);
            _scriptsService = new ScriptsService(unitOfWork);
        }

        private FiltroQuery montaFiltroQuery(Projeto projeto, string id, string valor, string operador = "")
        {
            FiltroQuery filtro = null;

            FiltroQuery.OperadorFiltro op = FiltroQuery.OperadorFiltro.Igual;
            if (operador == "0")
                op = FiltroQuery.OperadorFiltro.Igual;
            else if (operador == "1")
                op = FiltroQuery.OperadorFiltro.Contem;
            else if (operador == "2")
                op = FiltroQuery.OperadorFiltro.Diferente;
            else if (operador == "3")
                op = FiltroQuery.OperadorFiltro.MaiorQue;
            else if (operador == "4")
                op = FiltroQuery.OperadorFiltro.MenorQue;

            if (!id.Contains("L"))
            {
                if (id == "1")
                {
                    if (string.IsNullOrEmpty(operador))
                        op = FiltroQuery.OperadorFiltro.Contem;

                    filtro = new FiltroQuery
                    {
                        Tipo = FiltroQuery.TipoFiltro.Nome,
                        Operador = op,
                        Valor = valor
                    };
                }
                else if (id == "2")
                {
                    long idPasta = 0;
                    if (!Int64.TryParse(valor, out idPasta))
                    {
                        Pasta pasta = projeto.Pastas.FirstOrDefault(x => x.Nome == valor);
                        if (pasta != null)
                            idPasta = pasta.ID_Pasta;
                    }

                    filtro = new FiltroQuery
                    {
                        Tipo = FiltroQuery.TipoFiltro.Pasta,
                        Operador = op,
                        Valor = idPasta
                    };
                }
                else if (id == "3")
                {
                    if (string.IsNullOrEmpty(operador))
                        op = FiltroQuery.OperadorFiltro.Contem;

                    filtro = new FiltroQuery
                    {
                        Tipo = FiltroQuery.TipoFiltro.Titulo,
                        Operador = op,
                        Valor = valor
                    };
                }
                else if (id == "4")
                {
                    //Número Revisão
                }
                else if (id == "5")
                {
                    filtro = new FiltroQuery
                    {
                        Tipo = FiltroQuery.TipoFiltro.DataCriacao,
                        Operador = op,
                        Periodo = new PeriodoFiltro(valor, op)
                    };
                }
                else if (id == "6")
                {
                    //Criador Por
                }
                else if (id == "7")
                {
                    //Atualizado Por
                }
                else if (id == "8")
                {
                    //Responsável
                    Usuario usuario = _usuarioService.RetornarUsuario(valor);
                    if (usuario != null)
                    {
                        filtro = new FiltroQuery 
                        { 
                            Tipo = FiltroQuery.TipoFiltro.PendenciaFluxo,
                            Valor = usuario.ID_Usuario.ToString()
                        };
                    }
                }
                else if (id == "10")
                {
                    long idTipoItem = 0;
                    if (!Int64.TryParse(valor, out idTipoItem))
                    {
                        TipoItem tipoItem = projeto.TiposItem.FirstOrDefault(x => x.Nome == valor);
                        if (tipoItem != null)
                            idTipoItem = tipoItem.ID_TipoItem;
                    }

                    filtro = new FiltroQuery
                    {
                        Tipo = FiltroQuery.TipoFiltro.TipoItem,
                        Operador = op,
                        Valor = idTipoItem
                    };
                }
                else if (id == "16")
                {
                    long idBPMElemento = 0;
                    if (!Int64.TryParse(valor, out idBPMElemento))
                    {
                        foreach (Processo p in projeto.Processos)
                        {
                            BPMAtividade bpmAtividade = _bpmService.ListarAtividades(projeto).FirstOrDefault(a => (p.Nome + " - " + a.NomeTela) == valor);
                            if (bpmAtividade != null)
                            {
                                idBPMElemento = bpmAtividade.ID_BPMElemento;
                                break;
                            }
                        }
                    }

                    filtro = new FiltroQuery
                    {
                        Tipo = FiltroQuery.TipoFiltro.Atividade,
                        Operador = op,
                        Valor = idBPMElemento
                    };
                }
                else if (id == "25")
                {
                    filtro = new FiltroQuery
                    {
                        Tipo = FiltroQuery.TipoFiltro.DataCriacao,
                        Operador = op,
                        Periodo = new PeriodoFiltro(valor, op)
                    };
                }
                else if (id == "26")
                {
                    filtro = new FiltroQuery
                    {
                        Tipo = FiltroQuery.TipoFiltro.DataBaseRevisao,
                        Operador = op,
                        Periodo = new PeriodoFiltro(valor, op)
                    };
                }
                //todo? TipoFiltro.Conteudo;

                if (filtro != null && (filtro.Tipo == FiltroQuery.TipoFiltro.Nome || filtro.Tipo == FiltroQuery.TipoFiltro.Titulo || filtro.Tipo == FiltroQuery.TipoFiltro.Conteudo))
                {
                    string[] termos;
                    if (valor.Contains("\""))
                    {
                        termos = valor.Split('\"');
                    }
                    else
                        termos = valor.Split(' ');

                    termos = termos.Where(t => t.Length > 0).Select(t => t.Trim()).ToArray();

                    filtro.Valor = termos;
                }
            }
            else
            {
                long idMetadado = Convert.ToInt64(id.Substring(1));
                Metadado metadado = _projetoService.RetornarMetadado(idMetadado);

                if (metadado != null)
                {
                    switch (metadado.Tipo)
                    {
                        case "string":
                        case "memo":
                            if (string.IsNullOrEmpty(operador))
                                op = FiltroQuery.OperadorFiltro.Contem;

                            filtro = new FiltroQuery
                            {
                                ID_Metadado = metadado.ID_Metadado,
                                Tipo = FiltroQuery.TipoFiltro.Metadado,
                                Operador = op,
                                Valor = valor,
                                TipoMetadado = metadado.Tipo
                            };
                            break;
                        case "multi-valor":
                            //goto
                            break;
                        case "data":
                        case "data_hora":
                            filtro = new FiltroQuery(metadado, valor, 0, op);
                            break;
                        case "inteiro":
                        case "numero":
                        case "moeda":
                            filtro = new FiltroQuery
                            {
                                ID_Metadado = metadado.ID_Metadado,
                                Tipo = FiltroQuery.TipoFiltro.Metadado,
                                Operador = op,
                                Valor = valor,
                                TipoMetadado = metadado.Tipo
                            };
                            break;
                        case "lista":
                            switch (metadado.TipoLista)
                            {
                                case "unidades":
                                    var valorListaUnidades = metadado.Projeto.BPMUnidades.FirstOrDefault(lm => lm.Descricao == valor.ToString());
                                    if (valorListaUnidades != null)
                                    {
                                        filtro = new FiltroQuery
                                        {
                                            ID_Metadado = metadado.ID_Metadado,
                                            Tipo = FiltroQuery.TipoFiltro.Metadado,
                                            Operador = op,
                                            Valor = valorListaUnidades.ID_Unidades,
                                            TipoMetadado = metadado.Tipo
                                        };
                                    }
                                    else
                                    {
                                        var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                                        {
                                            Content = new StringContent("Incorrect Parameters, value " + valor + " not found.")
                                        };
                                        throw new HttpResponseException(resp);
                                    }
                                    break;
                                case "areas":
                                    var valorListaAreas = metadado.Projeto.BPMAreas.FirstOrDefault(lm => lm.Descricao == valor.ToString());
                                    if (valorListaAreas != null)
                                    {
                                        filtro = new FiltroQuery
                                        {
                                            ID_Metadado = metadado.ID_Metadado,
                                            Tipo = FiltroQuery.TipoFiltro.Metadado,
                                            Operador = op,
                                            Valor = valorListaAreas.ID_Area,
                                            TipoMetadado = metadado.Tipo
                                        };
                                    }
                                    else
                                    {
                                        var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                                        {
                                            Content = new StringContent("Incorrect Parameters, value " + valor + " not found.")
                                        };
                                        throw new HttpResponseException(resp);
                                    }
                                    break;
                                case "funcoes":
                                    var valorListaFuncoes = metadado.Projeto.BPMFuncoes.FirstOrDefault(lm => lm.Descricao == valor.ToString());
                                    if (valorListaFuncoes != null)
                                    {
                                        filtro = new FiltroQuery
                                        {
                                            ID_Metadado = metadado.ID_Metadado,
                                            Tipo = FiltroQuery.TipoFiltro.Metadado,
                                            Operador = op,
                                            Valor = valorListaFuncoes.ID_Funcoes,
                                            TipoMetadado = metadado.Tipo
                                        };
                                    }
                                    else
                                    {
                                        var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                                        {
                                            Content = new StringContent("Incorrect Parameters, value " + valor + " not found.")
                                        };
                                        throw new HttpResponseException(resp);
                                    }
                                    break;
                                case "usuarios":
                                    var valorListaUsuarios = metadado.Projeto.Conta.Usuarios.FirstOrDefault(u => u.Nome == valor.ToString());
                                    if (valorListaUsuarios != null)
                                    {
                                        filtro = new FiltroQuery
                                        {
                                            ID_Metadado = metadado.ID_Metadado,
                                            Tipo = FiltroQuery.TipoFiltro.Metadado,
                                            Operador = op,
                                            Valor = valorListaUsuarios.ID_Usuario,
                                            TipoMetadado = metadado.Tipo
                                        };
                                    }
                                    else
                                    {
                                        var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                                        {
                                            Content = new StringContent("Incorrect Parameters, value " + valor + " not found.")
                                        };
                                        throw new HttpResponseException(resp);
                                    }
                                    break;
                                case "grupos":
                                    if (metadado.Projeto.Grupos.Count == 0)
                                    {
                                        var valorListaGrupos = metadado.Projeto.Conta.Grupos.FirstOrDefault(g => g.Nome == valor.ToString());
                                        if (valorListaGrupos != null)
                                        {
                                            filtro = new FiltroQuery
                                            {
                                                ID_Metadado = metadado.ID_Metadado,
                                                Tipo = FiltroQuery.TipoFiltro.Metadado,
                                                Operador = op,
                                                Valor = valorListaGrupos.ID_Grupo,
                                                TipoMetadado = metadado.Tipo
                                            };
                                        }
                                        else
                                        {
                                            var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                                            {
                                                Content = new StringContent("Incorrect Parameters, value " + valor + " not found.")
                                            };
                                            throw new HttpResponseException(resp);
                                        }
                                    }
                                    else
                                    {
                                        List<Grupo> grupos = ProjetoHelper.GruposProjeto(metadado.Projeto, true);
                                        //List<Grupo> grupos = metadado.Projeto.Grupos.ToList();
                                        //grupos.AddRange(metadado.Projeto.GruposListas.ToList());
                                        //grupos = grupos.Distinct().OrderBy(g => g.Nome).ToList();

                                        var valorListaGrupos = grupos.FirstOrDefault(g => g.Nome == valor.ToString());
                                        if (valorListaGrupos != null)
                                        {
                                            filtro = new FiltroQuery
                                            {
                                                ID_Metadado = metadado.ID_Metadado,
                                                Tipo = FiltroQuery.TipoFiltro.Metadado,
                                                Operador = op,
                                                Valor = valorListaGrupos.ID_Grupo,
                                                TipoMetadado = metadado.Tipo
                                            };
                                        }
                                        else
                                        {
                                            var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                                            {
                                                Content = new StringContent("Incorrect Parameters, value " + valor + " not found.")
                                            };
                                            throw new HttpResponseException(resp);
                                        }
                                    }
                                    break;
                                case "valores_fixos":
                                    var valorListaValoresFixos = metadado.ListasMetadados.FirstOrDefault(lm => lm.Valor == valor.ToString());
                                    if (valorListaValoresFixos != null)
                                    {
                                        filtro = new FiltroQuery
                                        {
                                            ID_Metadado = metadado.ID_Metadado,
                                            Tipo = FiltroQuery.TipoFiltro.Metadado,
                                            Operador = op,
                                            Valor = valorListaValoresFixos.ID,
                                            TipoMetadado = metadado.Tipo
                                        };
                                    }
                                    else
                                    {
                                        var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                                        {
                                            Content = new StringContent("Incorrect Parameters, value " + valor + " not found.")
                                        };
                                        throw new HttpResponseException(resp);
                                    }
                                    break;
                            }
                            break;
                        case "lista_sigla":
                            var valoresListaSigla = metadado.ListasMetadados.FirstOrDefault(lm => lm.Valor == valor.ToString());

                            if (valoresListaSigla != null)
                            {
                                filtro = new FiltroQuery
                                {
                                    ID_Metadado = metadado.ID_Metadado,
                                    Tipo = FiltroQuery.TipoFiltro.Metadado,
                                    Operador = op,
                                    Valor = valoresListaSigla.ID,
                                    TipoMetadado = metadado.Tipo
                                };
                            }
                            else
                            {
                                valoresListaSigla = metadado.ListasMetadados.FirstOrDefault(lm => lm.Sigla == valor.ToString());

                                if (valoresListaSigla != null)
                                {
                                    filtro = new FiltroQuery
                                    {
                                        ID_Metadado = metadado.ID_Metadado,
                                        Tipo = FiltroQuery.TipoFiltro.Metadado,
                                        Operador = op,
                                        Valor = valoresListaSigla.ID,
                                        TipoMetadado = metadado.Tipo
                                    };
                                }
                                else
                                {
                                    valoresListaSigla = metadado.ListasMetadados.FirstOrDefault(lm => (lm.Sigla + " - " + lm.Valor) == valor.ToString());

                                    if (valoresListaSigla != null)
                                    {
                                        filtro = new FiltroQuery
                                        {
                                            ID_Metadado = metadado.ID_Metadado,
                                            Tipo = FiltroQuery.TipoFiltro.Metadado,
                                            Operador = op,
                                            Valor = valoresListaSigla.ID,
                                            TipoMetadado = metadado.Tipo
                                        };
                                    }
                                    else
                                    {
                                        var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                                        {
                                            Content = new StringContent("Incorrect Parameters, value " + valor + " not found.")
                                        };
                                        throw new HttpResponseException(resp);
                                    }
                                }
                            }

                            break;
                        case "checkbox":
                            filtro = new FiltroQuery
                            {
                                ID_Metadado = metadado.ID_Metadado,
                                Tipo = FiltroQuery.TipoFiltro.Metadado,
                                Operador = op,
                                Valor = valor,
                                TipoMetadado = metadado.Tipo
                            };
                            break;
                        case "local_Estrutura":
                            //goto
                            break;
                    }
                }
            }

            return filtro;
        }

        public JArray Get(long id, string tipo)
        {
            JArray result = new JArray();
            if (tipo == "tipoItem")
            {
                TipoItem tipoItem = _projetoService.RetornarTipoItem(id);
                foreach(var doc in tipoItem.Documentos)
                {
                    JObject docObj = new JObject();
                    docObj["ID"] = doc.ID_Documento;
                    docObj["Nome"] = doc.Nome;
                    docObj["TipoItem"] = tipoItem.Sigla + " - " +  tipoItem.Nome;
                    result.Add(docObj);
                }
            }
            else if (tipo == "ambiente")
            {
                Projeto projeto = _projetoService.RetornarProjeto(id);
                foreach (var tipoItem in projeto.TiposItem)
                {
                    foreach (var doc in tipoItem.Documentos)
                    {
                        JObject docObj = new JObject();
                        docObj["ID"] = doc.ID_Documento;
                        docObj["Nome"] = doc.Nome;
                        docObj["TipoItem"] = tipoItem.Sigla + " - " + tipoItem.Nome;
                        result.Add(docObj);
                    }
                }      
            }

            return result;
        }

        public JArray Get(long Projeto, string camposPesquisa, int retornaMetadados = 0, int retornaQuantidades = 0, bool todasRevisoes = false, long idTipoitem = 0, int linhaInicio = 0, int linhaFim = 0, string colunasRetorno = "")
        {
            Usuario usuario = _usuarioService.RetornarUsuario(Thread.CurrentPrincipal.Identity.Name);

            Projeto projeto = null;
            List<TipoItem> tiposItem = new List<TipoItem>();

            projeto = _projetoService.RetornarProjeto(Projeto);

            if(idTipoitem > 0)
            {
                var tipoItem = projeto.TiposItem.FirstOrDefault(ti => ti.ID_TipoItem == idTipoitem && ti.IncluirNaPesquisa);
                if(tipoItem != null)
                {
                    tiposItem.Add(tipoItem);
                }
            }
            else
            {
                tiposItem.AddRange(projeto.TiposItem.Where(t => t.IncluirNaPesquisa).OrderBy(t => t.Nome));
            }            

            JArray array = new JArray();

            if (tiposItem.Count > 0 && projeto != null)
            {
                List<FiltroQuery> filtros = new List<FiltroQuery>();

                #region Monta os Filtros

                Dictionary<string, string> camposFiltros = JsonConvert.DeserializeObject<Dictionary<string, string>>(camposPesquisa);

                foreach (KeyValuePair<string, string> d in camposFiltros)
                {
                    string nomeInterno = d.Key;
                    string valor = d.Value;

                    string id = "";
                    string operador = "";

                    Metadado metadado = _projetoService.RetornarMetadado(projeto.ID_Projeto, nomeInterno);
                    if(metadado != null)
                    {
                        id = "L" + metadado.ID_Metadado;
                    }
                    else
                    {
                        //Caso não tenha sido enviado o nome_interno, provavelmente foi enviado o id
                        string[] separadorT = new string[] { "|_|" };
                        string[] termosAvancados = d.Value.Split(separadorT, StringSplitOptions.None);

                        id = nomeInterno;
                        operador = termosAvancados[0];
                        valor = termosAvancados[1];
                    }

                    FiltroQuery fq = montaFiltroQuery(projeto, id, valor, operador);

                    if (fq != null)
                    {
                        filtros.Add(fq);
                    }
                }

                #endregion

                bool exportacaoDados = usuario.Services == true || projeto.PermissoesEspeciais.Count(x => x.PermiteExportacaoPlanilha == true && x.ID_Usuario == usuario.ID_Usuario) > 0;

                if (retornaQuantidades > 0)
                {
                    var agrupamentosTipoItem = _itensService.QueryItensAgrupadosTotalArquivos(projeto, usuario,
                        tiposItem, filtros, "ID_TipoItem", todasRevisoes, exportacaoDados);

                    array = RetornoPesquisaQuantidade(agrupamentosTipoItem);
                }
                else
                {
                    List<ColunaQuery> colunas = new List<ColunaQuery>();
                    colunas.Add(new ColunaQuery(ColunaQuery.TipoColuna.Revisao));
                    colunas.Add(new ColunaQuery(ColunaQuery.TipoColuna.DataCriacao));
                    colunas.Add(new ColunaQuery(ColunaQuery.TipoColuna.DataAtualizacao));
                    colunas.Add(new ColunaQuery(ColunaQuery.TipoColuna.DataRevisao));
                    colunas.Add(new ColunaQuery(ColunaQuery.TipoColuna.Situacao));
                    //colunas.Add(new ColunaQuery(ColunaQuery.TipoColuna.Responsavel));

                    var colunasDescricao = ColunaQuery.ColunasDescricao(tiposItem);
                    colunasDescricao.Add(new ColunaQuery(ColunaQuery.TipoColuna.ColunaDescricao));
                    colunas.AddRange(colunasDescricao);

                    List<Metadado> metadados = new List<Metadado>();

                    if (!string.IsNullOrEmpty(colunasRetorno))
                    {
                        colunasRetorno = colunasRetorno.Replace("\"", "");
                        string[] _colunsNomesInternos = colunasRetorno.Split(',');
                        foreach (string _col in _colunsNomesInternos)
                        {
                            Metadado metadado = _projetoService.RetornarMetadado(projeto.ID_Projeto, _col);

                            if (metadado != null)
                            {
                                metadados.Add(metadado);
                            }
                        }

                        foreach (var metadado in metadados.Distinct())
                        {
                            if (colunas.Any(c => c.ID_Metadado == metadado.ID_Metadado))
                            {
                                continue;
                            }

                            colunas.Add(new ColunaQuery(metadado));
                        }
                    }
                    else
                    {
                        if (retornaMetadados > 0)
                        {
                            List<Metadado> tabelas = new List<Metadado>();
                            foreach (var item in projeto.TiposItem.Where(t => t.PropriedadesTipoItem.Count(p => p.ID_Metadado != null && p.Metadado.Tipo == "tabela") > 0))
                            {
                                tabelas.AddRange(item.PropriedadesTipoItem.Where(p => p.ID_Metadado != null && p.Metadado.Tipo == "tabela").Select(p => p.Metadado).ToList());
                            }

                            foreach (var tipoItem in tiposItem)
                            {
                                metadados.AddRange(tipoItem.PropriedadesTipoItem.Where(x => x.ID_Metadado.HasValue && tabelas.Count(t => t.ColunasMetadado.Count(cm => cm.ID_MetadadoColuna == x.ID_Metadado) > 0) == 0).OrderBy(x => x.Ordem).Select(x => x.Metadado));
                            }

                            foreach (var metadado in metadados.Distinct())
                            {
                                colunas.Add(new ColunaQuery(metadado));
                            }
                        }
                    }

                    foreach (var tipoItem in tiposItem)
                    {
                        foreach (var processo in tipoItem.Processos)
                        {
                            var atividadesDistribuicao = processo.BPMElementos.OfType<BPMAtividade>().Where(a => a.TipoAtividade == 5);
                            foreach (var atividade in atividadesDistribuicao)
                            {
                                filtros.Add(new FiltroQuery { Tipo = FiltroQuery.TipoFiltro.AtividadeFixo, Operador = FiltroQuery.OperadorFiltro.Diferente, Valor = atividade.ID_BPMElemento });
                            }
                        }
                    }


                    bool colunasJson = _projetoService.RetornaValorConstante(projeto.ID_Projeto, "CAMPOS_JSON") == 1;
                    
                    if (linhaInicio == 0 && exportacaoDados) 
                    {
                        var agrupamentosTipoItemCheck = _itensService.QueryItensAgrupadosTotalArquivos(projeto, usuario,
                        tiposItem, filtros, "ID_TipoItem", todasRevisoes, exportacaoDados);

                        int somaTiposItem = agrupamentosTipoItemCheck.Sum(x => x.Total);

                        int checkCount = _itensService.QueryItens(projeto, usuario, tiposItem, filtros, colunas, new OrdemQuery(), linhaInicio, linhaFim, todasRevisoes, exportacaoDados: exportacaoDados, colunasJson: colunasJson, apenasContar: true);

                        if (checkCount != somaTiposItem)
                        {
                            //retorna um erro.
                            throw new Exception("Divergência na quantidade de documentos, verifique se não existem valores de metadados duplicados no banco de dados.");
                        }
                    }

                    var docs = _itensService.QueryItens(projeto, usuario, tiposItem, filtros, colunas, new OrdemQuery(), linhaInicio, linhaFim, todasRevisoes, exportacaoDados: exportacaoDados, colunasJson:colunasJson);

                    array = RetornoPesquisa(docs, metadados);
                }
            }

            return array;
        }

        public JArray Get(string pesquisa, int retornaMetadados = 0, long Projeto = 0, long TipoItem = 0)
        {
            Usuario usuario = _usuarioService.RetornarUsuario(Thread.CurrentPrincipal.Identity.Name);

            Projeto projeto = null;
            List<TipoItem> tiposItem = new List<TipoItem>();

            if (Projeto > 0)
            {
                projeto = _projetoService.RetornarProjeto(Projeto);

                if (TipoItem > 0)
                {
                    TipoItem tipoItem = projeto.TiposItem.FirstOrDefault(x => x.ID_TipoItem == TipoItem);
                    if (tipoItem != null)
                        tiposItem.Add(tipoItem);
                }
                else
                {
                    tiposItem.AddRange(projeto.TiposItem.Where(t => t.IncluirNaPesquisa).OrderBy(t => t.Nome));
                }
            }
            else
            {
                List<Projeto> projetos = usuario.Conta.Projetos.ToList();

                foreach (Projeto proj in projetos)
                {
                    if (TipoItem > 0)
                    {
                        TipoItem tipoItem = proj.TiposItem.FirstOrDefault(x => x.ID_TipoItem == TipoItem);
                        if (tipoItem != null)
                        {
                            tiposItem.Add(tipoItem);
                            break;
                        }
                    }
                    else
                    {
                        tiposItem.AddRange(proj.TiposItem.Where(t => t.IncluirNaPesquisa).OrderBy(t => t.Nome));
                    }
                }

                if (tiposItem.Count > 0)
                {
                    projeto = tiposItem.First().Projeto;
                }
            }

            JArray array = new JArray();

            string[] termo;
            termo = pesquisa.Split(' ');

            termo = termo.Where(t => t.Length > 0).Select(t => t.Trim()).ToArray();

            if (termo.Length > 0 && tiposItem.Count > 0 && projeto != null)
            {
                List<FiltroQuery> filtros = new List<FiltroQuery>();
                filtros.Add(new FiltroQuery
                {
                    Tipo = FiltroQuery.TipoFiltro.Indice,
                    Operador = FiltroQuery.OperadorFiltro.Contem,
                    Valor = termo
                });

                var colunasDescricao = ColunaQuery.ColunasDescricao(tiposItem);
                colunasDescricao.Add(new ColunaQuery(ColunaQuery.TipoColuna.ColunaDescricao));

                List<ColunaQuery> colunas = new List<ColunaQuery>();
                colunas.AddRange(colunasDescricao);

                List<Metadado> metadados = new List<Metadado>();
                if (retornaMetadados > 0)
                {
                    foreach (var tipoItem in tiposItem)
                    {
                        metadados.AddRange(tipoItem.PropriedadesTipoItem.Where(x => x.ID_Metadado.HasValue).Where(x => x.Metadado.Tipo == "string" || x.Metadado.Tipo == "data" || x.Metadado.Tipo == "inteiro" || x.Metadado.Tipo == "lista" || x.Metadado.Tipo == "lista_sigla" || x.Metadado.Tipo == "calculado" || x.Metadado.Tipo == "numero" || x.Metadado.Tipo == "moeda" || x.Metadado.Tipo == "multi-valor" || x.Metadado.Tipo == "data_hora" || x.Metadado.Tipo == "memo" || x.Metadado.Tipo == "local_Estrutura" || x.Metadado.Tipo == "form_relacionado").OrderBy(x => x.Ordem).Select(x => x.Metadado));
                    }

                    foreach (var metadado in metadados.Distinct())
                    {
                        colunas.Add(new ColunaQuery(metadado));
                    }
                }

                bool colunasJson = _projetoService.RetornaValorConstante(projeto.ID_Projeto, "CAMPOS_JSON") == 1;

                var docs = _itensService.QueryItens(projeto, usuario, tiposItem, filtros, colunas, new OrdemQuery(), 0, 0, colunasJson: colunasJson);

                array = RetornoPesquisa(docs, metadados);
            }

            return array;
        }

        private JArray RetornoPesquisa(dynamic docs, List<Metadado> metadados)
        {
            JArray array = new JArray();

            foreach (var item in docs)
            {
                JObject objeto = new JObject();

                objeto["ID_Documento"] = item.ID_Documento;
                objeto["Nome_Documento"] = item.Nome;
                objeto["ID_TipoItem"] = item.ID_TipoItem;
                objeto["ID_RevisaoAtual"] = item.ID_RevisaoAtual;
                objeto["NumeroRevisao"] = item.NumeroRevisao;
                objeto["Descricao"] = item.Descricao;
                objeto["DataCriacao"] = item.DataCriacao;
                objeto["DataAtualizacao"] = item.DataAtualizacao;
                objeto["DataRevisao"] = item.DataRevisao;
                objeto["NomeAtividade"] = item.NomeAtividade;
                objeto["Situacao"] = item.Situacao;
                objeto["Responsavel"] = item.Responsavel;

                Dictionary<string, string> valoresColunasJson = new Dictionary<string, string>();
                if (item.Campos != null)
                    valoresColunasJson = JsonConvert.DeserializeObject<Dictionary<string, string>>(item.Campos);

                foreach (var metadado in metadados)
                {
                    string nomeCol = "c" + metadado.ID_Metadado;

                    string valor = "";
                    
                    if (item.GetType().GetProperty(nomeCol) != null)
                    {
                        valor = item.GetType().GetProperty(nomeCol).GetValue(item, null);
                        valor = valor != null ? valor : "";
                    }
                    else
                    {
                        if (valoresColunasJson.ContainsKey(nomeCol))
                        {
                            valor = valoresColunasJson[nomeCol];
                        }
                    }

                    objeto[metadado.NomeInterno] = valor;
                }

                objeto["ColunasMetadados"] = string.Join(";", metadados.Select(m => m.NomeInterno).ToArray());

                array.Add(objeto);
            }

            return array;
        }

        private JArray RetornoPesquisaQuantidade(List<AgrupamentoQuery> agrupamentosTipoItem)
        {
            JArray array = new JArray();

            foreach (var item in agrupamentosTipoItem.Where(x => x.Chave.HasValue))
            {
                JObject objeto = new JObject();

                TipoItem tipoItemAgrupamento = _projetoService.RetornarTipoItem(item.Chave.Value);
                if(tipoItemAgrupamento != null)
                {
                    objeto["ID_TipoItem"] = item.Chave.Value;
                    objeto["TipoItem"] = tipoItemAgrupamento.Nome;
                    objeto["QuantidadeItens"] = item.Total;
                    objeto["QuantidadeArquivos"] = item.TotalArquivos;
                }

                array.Add(objeto);
            }

            return array;
        }

        public JObject Get(long ID_Revisao)
        {
            JObject objeto = null;

            Revisao revisao = _documentoService.RetornaRevisao(ID_Revisao);
            Usuario usuario = _usuarioService.RetornarUsuario(Thread.CurrentPrincipal.Identity.Name);

            if (revisao != null)
            {
                Documento documento = revisao.Documento;
                BPMInstancia instancia = _bpmService.RetornaInstanciaResponsavel(usuario.ID_Usuario, documento.ID_Documento);

                if (instancia == null || (instancia != null && _documentoService.VerificaAcessoInstancia(instancia, usuario)))
                {
                    objeto = new JObject();

                    objeto["ID_Documento"] = documento.ID_Documento;
                    objeto["Nome_Documento"] = revisao.NomeDocumento;
                    objeto["Natureza"] = documento.TipoItem.Natureza;
                    objeto["Descricao"] = _documentoService.ColunaDescricao(documento);
                    objeto["Numero_Revisao"] = revisao.Numero;
                    objeto["ID_Revisao"] = revisao.ID_Revisao;
                    objeto["Nome_Projeto"] = documento.Projeto.Nome;
                    objeto["ID_Processo"] = documento.ID_ProcessoOrigem;
                    objeto["ID_TipoItem"] = documento.ID_TipoItem;

                    bool usuarioResponsavel = _documentoService.VerificaResponsabilidade(documento, usuario, instancia.ID_BPMInstancia) && !instancia.Concluido;
                    bool executarQualquerAcaoFluxo = _contaService.UsuarioPossuiPermissao(documento.Projeto, usuario, Usuario.Permissoes.ExecutarQualquerAcaoFluxo);
                    bool adminProcesso = _processoService.UsuarioAdminProcesso(usuario, documento.ProcessoOrigem, documento);

                    objeto["Editavel"] = usuarioResponsavel;

                    bool podeAtualizarArquivo = false;
                    if (usuarioResponsavel)
                    {
                        podeAtualizarArquivo = _arquivoService.VerificaPermissaoParaAtualizarArquivo(documento, instancia, null, usuario);
                    }

                    objeto["PodeAtualizarArquivo"] = podeAtualizarArquivo;
                    objeto["Multiarquivo"] = documento.TipoItem.PermiteMultiArquivo;

                    bool utilizaRevisaoLiberada = false;

                    foreach (var processo in documento.TipoItem.Processos)
                    {
                        if (_processoService.PossuiEventoRevisaoLiberada(processo))
                        {
                            utilizaRevisaoLiberada = true;
                            break;
                        }
                    }

                    JArray arrayAcoes = new JArray();
                    if (usuarioResponsavel || executarQualquerAcaoFluxo || adminProcesso)
                    {
                        List<BPMAcao> acoes = _bpmService.AcoesInstancia(instancia.ID_BPMInstancia, usuario);

                        foreach (BPMAcao acao in acoes.Where(a => a.TipoAcao != 8))
                        {
                            JObject objAcao = new JObject();
                            objAcao["NomeInterno"] = acao.NomeInterno;
                            objAcao["Nome"] = acao.Nome;
                            objAcao["Tipo"] = acao.TipoAcao;

                            arrayAcoes.Add(objAcao);
                        }
                    }
                    objeto["Acoes"] = arrayAcoes;

                    if (revisao.ID_Revisao == documento.ID_RevisaoAtual && (revisao.ID_Revisao == documento.ID_RevisaoLiberada || !utilizaRevisaoLiberada))
                    {
                        objeto["Situacao"] = "Vigente";
                    }
                    else if ((revisao.ID_Revisao != documento.ID_RevisaoAtual && revisao.ID_Revisao == documento.ID_RevisaoLiberada) ||
                            (revisao.ID_Revisao == documento.ID_RevisaoAtual && documento.ID_RevisaoLiberada != documento.ID_RevisaoAtual))
                    {
                        objeto["Situacao"] = "Em Revisão";
                    }
                    else if (revisao.ID_Revisao != documento.ID_RevisaoLiberada)
                    {
                        objeto["Situacao"] = "Obsoleto";
                    }

                    objeto["ID_RevisaoLiberada"] = documento.ID_RevisaoLiberada;

                    objeto["ID_Arquivo"] = 0;
                    objeto["Nome_Arquivo"] = "";
                    Arquivo arquivo = revisao.Arquivos.FirstOrDefault(a => a.Extensao.ToLower() == ".pdf");
                    if (arquivo == null)
                    {
                        arquivo = revisao.Arquivos.FirstOrDefault(a => Utils.ExtensoesArquivos().Any(e => e == a.Extensao.ToUpper()));
                    }

                    if (arquivo != null)
                    {
                        objeto["ID_Arquivo"] = arquivo.ID_Arquivo;
                        objeto["Nome_Arquivo"] = arquivo.Nome + arquivo.Extensao;
                    }
                    List<Arquivo> listaArquivos = new List<Arquivo>();
                    listaArquivos = _arquivoService.ListarArquivos(revisao);

                    JArray arrayListaArquivos = new JArray();
                    foreach (Arquivo arquivoVez in listaArquivos)
                    {
                        JObject objArquivo = new JObject();
                        objArquivo["PodeBaixar"] = _arquivoService.PodeBaixarArquivo(instancia, false, revisao.ID_Revisao, usuario);
                        objArquivo["PodeVisualizar"] = false;
                        if (arquivoVez.Extensao.ToUpper() == ".DWG" || arquivoVez.Extensao.ToUpper() == ".DGN" || arquivoVez.Extensao.ToUpper() == ".DXF" || 
                            arquivoVez.Extensao.ToUpper() == ".MPP" ||
                            arquivoVez.Extensao.ToUpper() == ".DOC" || arquivoVez.Extensao.ToUpper() == ".DOCX" ||
                            arquivoVez.Extensao.ToUpper() == ".XLS" || arquivoVez.Extensao.ToUpper() == ".XLSX" ||
                            arquivoVez.Extensao.ToUpper() == ".PPT" || arquivoVez.Extensao.ToUpper() == ".PPTX" || 
                            arquivoVez.Extensao.ToUpper() == ".PDF" || arquivoVez.Extensao.ToUpper() == ".TIF" ||
                            arquivoVez.Extensao.ToUpper() == ".TIFF" || arquivoVez.Extensao.ToUpper() == ".PNG" ||
                            arquivoVez.Extensao.ToUpper() == ".JPEG" || arquivoVez.Extensao.ToUpper() == ".JPG")
                            objArquivo["PodeVisualizar"] = true;
                        objArquivo["ID_Arquivo"] = arquivoVez.ID_Arquivo;
                        objArquivo["Nome_Arquivo"] = arquivoVez.Nome + arquivoVez.Extensao;
                        objArquivo["Extensao"] = arquivoVez.Extensao;
                        objArquivo["UsuarioQueIncluiu"] = arquivoVez.Usuario.Nome;
                        objArquivo["Tamanho"] = arquivoVez.Tamanho;
                        objArquivo["DataInclusao"] = arquivoVez.Data;

                        objArquivo["PodeAtualizar"] = _arquivoService.VerificaPermissaoParaAtualizarArquivo(documento, instancia, arquivoVez, usuario);

                        arrayListaArquivos.Add(objArquivo);
                    }

                    objeto["ListaArquivos"] = arrayListaArquivos;
                    BPMElemento elementoAtual = null;

                    if (instancia != null && instancia.ID_BPMElemento > 0)
                    {
                        elementoAtual = _bpmService.RetornarBPMElemento(Convert.ToInt64(instancia.ID_BPMElemento));
                    }

                    objeto["Atividade"] = elementoAtual.NomeTela;

                    bool fimDeFluxo = false;
                    if (elementoAtual is BPMEvento)
                    {
                        BPMEvento eventoAtual = elementoAtual as BPMEvento;
                        fimDeFluxo = eventoAtual.TipoEvento == (int)BPMEvento.TiposEvento.FimFinal;
                    }
                    else if (elementoAtual is BPMAtividade)
                    {
                        BPMAtividade atividadeAtual = elementoAtual as BPMAtividade;
                        fimDeFluxo = atividadeAtual.FimDeFluxo;
                    }

                    List<PropriedadeTipoItem> propriedades = _projetoService.ListarCamposConjuntosOrdenados(documento.TipoItem, elementoAtual);

                    var idsMetadadosTabela = propriedades.Where(p => p.ID_Metadado > 0 && p.Metadado.Tipo == "tabela").Select(p => p.ID_Metadado);

                    List<PropriedadeTipoItem> propsSemColunasTabela = propriedades.Where(p => (
                        (p.ID_Metadado == null && p.ID_ConjuntoMetadado != null && p.ConjuntoMetadado.Visivel) ||
                        (p.ID_Metadado > 0 && p.Metadado.TabelasMetadado.Count(tb => idsMetadadosTabela.Any(idtab => idtab == tb.ID_MetadadoTabela)) == 0)
                    )).ToList();

                    List<MetadadosRegra> regras = new List<MetadadosRegra>();

                    JArray valores = new JArray();

                    foreach (var prop in propsSemColunasTabela)
                    {
                        JObject valorMetadado = new JObject();

                        if (prop.ID_Metadado > 0)
                        {
                            bool somenteLeitura = prop.SomenteLeitura;
                            if (fimDeFluxo)
                            {
                                somenteLeitura = true;
                            }

                            if (usuario.PermissoesEspeciais.Any(c => c.EditarCampoSomenteLeitura))
                                somenteLeitura = false;

                            valorMetadado["ID"] = prop.ID_Metadado;
                            valorMetadado["Nome"] = prop.Metadado.Nome;
                            if (prop.Metadado.Tipo == "data_hora" || prop.Metadado.Tipo == "data")
                            {
                                ValorMetadado vm = revisao.ValoresMetadados.FirstOrDefault(vms => vms.ID_Metadado == prop.ID_Metadado);
                                if (vm != null)
                                    valorMetadado["Valor"] = vm.Data;
                            }
                            else if (prop.Metadado.Tipo == "form_relacionado")
                            {
                                string valor = _documentoService.RetornarValorMetadadoRevisao(revisao, prop.Metadado);
                                MetadadoTiposItensRelacionados mtr = prop.Metadado.TiposItensRelacionados.FirstOrDefault();
                                if (mtr != null)
                                {
                                    valorMetadado["ID_TipoItemRel"] = mtr.ID_TipoItemRelacionado;
                                    valorMetadado["ID_ProjetoRel"] = mtr.TipoItem.ID_Projeto;
                                    if (prop.Metadado.RelacionarAoMetadado)
                                    {
                                        if (!string.IsNullOrEmpty(valor))
                                        {
                                            string[] idsRefs = valor.Split(',');

                                            List<string> nomesRefs = new List<string>();

                                            foreach (var id in idsRefs)
                                            {
                                                long idRef = 0;
                                                Int64.TryParse(id, out idRef);
                                                if (idRef > 0)
                                                {
                                                    Referencia referencia = _unitOfWork.Referencias.GetByID(idRef);
                                                    if (mtr.SentidoRelacionamento == 1)
                                                        nomesRefs.Add(referencia.DocumentosReferencia.Nome);
                                                    else
                                                        nomesRefs.Add(referencia.Documentos.Nome);
                                                }
                                            }

                                            valor = string.Join(",", nomesRefs);
                                        }
                                    }
                                }
                                valorMetadado["Valor"] = valor;
                            }
                            else if (prop.Metadado.Tipo == "arquivo")
                            {
                                string valor = _documentoService.RetornarValorMetadadoRevisao(revisao, prop.Metadado);
                                string nomeArquivo = "";
                                long idArquivo = 0;
                                Int64.TryParse(valor, out idArquivo);
                                if (idArquivo > 0)
                                {
                                    Arquivo arqCampo = _arquivoService.RetornarArquivo(idArquivo);
                                    if (arqCampo != null)
                                    {
                                        nomeArquivo = arqCampo.Nome + arqCampo.Extensao;
                                    }
                                }

                                valorMetadado["Valor"] = valor;
                                valorMetadado["NomeArquivo"] = nomeArquivo;
                            }
                            else
                            {
                                valorMetadado["Valor"] = _documentoService.RetornarValorMetadadoRevisao(revisao, prop.Metadado);
                            }
                            valorMetadado["Tipo"] = prop.Metadado.Tipo;
                            valorMetadado["TipoLista"] = prop.Metadado.TipoLista != null ? prop.Metadado.TipoLista : "";
                            valorMetadado["TamanhoTela"] = prop.Metadado.TamanhoTela != null ? prop.Metadado.TamanhoTela : "";
                            valorMetadado["Obrigatorio"] = prop.Obrigatorio;
                            valorMetadado["SomenteLeitura"] = somenteLeitura;
                            valorMetadado["Metadado"] = true;
                            valorMetadado["ID_MetadadoPai"] = prop.Metadado.ID_MetadadoCascata;
                            valorMetadado["ExibeSigla"] = prop.Metadado.ExibeSigla;

                            if (prop.Metadado.Tipo == "label")
                            {
                                valorMetadado["Metadado"] = false;
                            }
                            else if (prop.Metadado.Tipo == "tabela")
                            {
                                var linhas = new JArray();

                                //Cabeçalho da tabela
                                var linha = new JObject();
                                linha["ID_Linha"] = 0;

                                var colunas = new JArray();
                                foreach (var coluna in prop.Metadado.ColunasMetadado)
                                {
                                    JObject header = new JObject();
                                    header["ID_Metadado"] = coluna.ID_MetadadoColuna;
                                    header["Value"] = coluna.Coluna.Nome;
                                    header["Tipo"] = coluna.Coluna.Tipo;
                                    header["TipoLista"] = coluna.Coluna.TipoLista != null ? coluna.Coluna.TipoLista : "";
                                    header["NomeInterno"] = coluna.Coluna.NomeInterno;
                                    header["ID_MetadadoPai"] = coluna.Coluna.ID_MetadadoCascata;
                                    header["ExibeSigla"] = coluna.Coluna.ExibeSigla;

                                    PropriedadeTipoItem propColuna = propriedades.FirstOrDefault(x => x.ID_Metadado == coluna.ID_MetadadoColuna);
                                    if (propColuna != null)
                                    {
                                        bool colSomenteLeitura = propColuna.SomenteLeitura;
                                        if (fimDeFluxo)
                                            colSomenteLeitura = true;

                                        header["Obrigatorio"] = propColuna.Obrigatorio;
                                        header["SomenteLeitura"] = colSomenteLeitura;
                                    }

                                    colunas.Add(header);
                                }

                                linha["Colunas"] = colunas;
                                linhas.Add(linha);

                                //Corpo da tabela
                                foreach (var vt in revisao.ValoresTabelas.Where(t => t.ID_Metadado == prop.ID_Metadado))
                                {
                                    linha = new JObject();
                                    linha["ID_Linha"] = vt.ID_ValorTabela;

                                    colunas = new JArray();
                                    foreach (var coluna in prop.Metadado.ColunasMetadado)
                                    {
                                        JToken valor;

                                        ValorMetadado vm = vt.ValoresMetadados.FirstOrDefault(v => v.ID_Metadado == coluna.ID_MetadadoColuna);
                                        if (vm != null)
                                        {
                                            if (vm.Metadado.Tipo == "data_hora" || vm.Metadado.Tipo == "data")
                                                valor = vm.Data;
                                            else
                                                valor = _documentoService.RetornarValorMetadadoDocumento(vm);
                                        }
                                        else
                                        {
                                            valor = "";
                                        }

                                        JObject body = new JObject();
                                        body["ID_Metadado"] = coluna.ID_MetadadoColuna;
                                        body["Value"] = valor;
                                        if (coluna.Coluna.Tipo == "arquivo") 
                                        {
                                            string nomeArquivo = "";
                                            long idArquivo = 0;
                                            Int64.TryParse(valor.ToString(), out idArquivo);
                                            if (idArquivo > 0)
                                            {
                                                Arquivo arqCampo = _arquivoService.RetornarArquivo(idArquivo);
                                                if (arqCampo != null)
                                                {
                                                    nomeArquivo = arqCampo.Nome + arqCampo.Extensao;
                                                }
                                            }

                                            body["NomeArquivo"] = nomeArquivo;
                                        }

                                        colunas.Add(body);
                                    }

                                    linha["Colunas"] = colunas;
                                    linhas.Add(linha);
                                }

                                valorMetadado["ValorTabela"] = linhas;

                                valorMetadado["ExpandirLinhas"] = prop.Metadado.ExpandirLinhasTabelaApp;
                                valorMetadado["ColunaExpandida"] = prop.Metadado.ColunaExpandidaTabelaApp;
                            }

                            regras.AddRange(prop.Metadado.Regras.ToList());
                        }
                        else if (prop.ID_ConjuntoMetadado != null)
                        {
                            valorMetadado["ID"] = prop.ID_ConjuntoMetadado;
                            valorMetadado["Nome"] = prop.ConjuntoMetadado.Nome;
                            valorMetadado["Valor"] = "";
                            valorMetadado["Tipo"] = "";
                            valorMetadado["TipoLista"] = "";
                            valorMetadado["TamanhoTela"] = "";
                            valorMetadado["Obrigatorio"] = false;
                            valorMetadado["SomenteLeitura"] = false;
                            valorMetadado["Metadado"] = false;
                        }

                        valores.Add(valorMetadado);
                    }

                    objeto["Valores_Metadados"] = valores;

                    JArray arrayRegras = new JArray();

                    foreach (var item in regras)
                    {
                        JObject objRegra = new JObject();
                        objRegra["ID_Metadado"] = item.ID_Metadado;
                        objRegra["ID_MetadadoRegra"] = item.ID_MetadadoRegra;
                        objRegra["TipoRegra"] = item.TipoRegra;
                        objRegra["Valor"] = item.Valor;

                        arrayRegras.Add(objRegra);
                    }

                    objeto["Regras"] = arrayRegras;
                }
            }

            return objeto;
        }

        // GET: api/Itens/5
        public JObject Get()
        {
            string NomeDocumento = HttpContext.Current.Request.Params["NomeDocumento"];
            string TipoItem = HttpContext.Current.Request.Params["TipoItem"];
            string login = HttpContext.Current.Request.Params["LoginUsuario"];
            string ID_Documento = HttpContext.Current.Request.Params["ID_Documento"];

            if (string.IsNullOrEmpty(NomeDocumento) && string.IsNullOrEmpty(ID_Documento))
            {
                var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                {
                    Content = new StringContent("Incorrect Parameters")
                };
                throw new HttpResponseException(resp);

            }

            List<long> idsTipoItem = new List<long>();

            if (!string.IsNullOrEmpty(TipoItem))
            {
                foreach (var item in TipoItem.Split(','))
                {
                    long id_tipoitem = 0;
                    if (Int64.TryParse(item, out id_tipoitem))
                        idsTipoItem.Add(id_tipoitem);
                }
            }
            
            Documento d = null;

            if (!string.IsNullOrEmpty(ID_Documento))
            {
                long idDocumento = 0;
                if (Int64.TryParse(ID_Documento, out idDocumento))
                {
                    d = _documentoService.RetornarDocumento(Convert.ToInt64(ID_Documento));

                    if (d == null)
                    {
                        var resp = new HttpResponseMessage(HttpStatusCode.NotFound)
                        {
                            Content = new StringContent(string.Format("No document with ID = {0}", idDocumento)),
                            ReasonPhrase = "Document ID Not Found"
                        };
                        throw new HttpResponseException(resp);
                    }
                }
                else
                {
                    var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                    {
                        Content = new StringContent("Parameter was in incorrect format: ID_Documento")
                    };
                    throw new HttpResponseException(resp);

                }
            }
            else
            {
                foreach (var id_tipoitem in idsTipoItem.Distinct())
                {
                    TipoItem tipoItem = _projetoService.RetornarTipoItem(id_tipoitem);
                    Projeto projeto = tipoItem.Projeto;

                    _documentoService.projetoSelecionado = projeto;
                    d = _documentoService.RetornarDocumento(NomeDocumento, projeto.ID_Projeto, tipoItem.ID_TipoItem, false);

                    if (d != null)
                        break;
                }

                if (d == null)
                {
                    var resp = new HttpResponseMessage(HttpStatusCode.NotFound)
                    {
                        Content = new StringContent(string.Format("No document with Name = {0}", NomeDocumento)),
                        ReasonPhrase = "Document Not Found"
                    };
                    throw new HttpResponseException(resp);
                }
            }

            Usuario usuario = null;
            if (login != null && login != "")
            {
                usuario = _usuarioService.RetornarUsuario(login);
            }

            JObject objeto = new JObject();

            if (d != null)
            {
                objeto = montaJSON(d, 1, usuario);
            }

            return objeto;
        }

        private JObject montaJSON(Documento d, int retornaMetadados = 0, Usuario usuario = null)
        {
            JObject objeto = new JObject();

            objeto["ID_Documento"] = d.ID_Documento;
            objeto["Nome_Documento"] = d.Nome;
            objeto["ID_TipoItem"] = d.ID_TipoItem;

            objeto["ID_Revisao"] = d.ID_RevisaoAtual;
            objeto["ID_RevisaoLiberada"] = d.ID_RevisaoLiberada;

            BPMInstancia instancia = _bpmService.RetornaInstanciaItem(d.ID_Documento);
            if (instancia != null)
                objeto["NomeAtividade"] = instancia.BPMElemento.NomeTela;

            if (retornaMetadados > 0)
            {
                foreach (PropriedadeTipoItem metadado in d.TipoItem.PropriedadesTipoItem.Where(p => p.Metadado != null))
                {
                    string v = "";

                    ValorMetadado valorMetadado = d.RevisaoAtual.ValoresMetadados.FirstOrDefault(vm => vm.ID_Metadado == metadado.ID_Metadado);

                    if (valorMetadado != null)
                    {
                        switch (metadado.Metadado.Tipo)
                        {
                            case "string":
                            case "memo":
                                v = valorMetadado.String;
                                break;
                            case "multi-valor":
                                v = _documentoService.RetornarValorMetadadoDocumento(valorMetadado);
                                break;
                            case "data":
                            case "data_hora":
                                v = valorMetadado.Data.ToString();
                                break;
                            case "inteiro":
                            case "numero":
                            case "moeda":
                                v = valorMetadado.Inteiro.ToString();
                                break;
                            case "lista":
                                v = _documentoService.RetornarValorMetadadoDocumento(valorMetadado);
                                break;
                            case "lista_sigla":
                                v = _documentoService.RetornarValorMetadadoDocumento(valorMetadado);
                                break;
                            case "checkbox":
                                v = valorMetadado.Bool.ToString();
                                break;
                            case "local_Estrutura":
                                if (valorMetadado.Inteiro.HasValue)
                                {
                                    Pasta pasta = _pastaService.RetornarPasta((long)valorMetadado.Inteiro);
                                    if (pasta.Sigla != "")
                                    {
                                        v = pasta.Sigla + "-" + pasta.Nome;
                                    }
                                    else
                                    {
                                        v = pasta.Nome;
                                    }
                                }
                                break;
                        }

                        objeto[valorMetadado.Metadado.NomeInterno] = v;
                    }

                    if (metadado.Metadado.Tipo == "tabela") //copia tabelas
                    {
                        #region campos tabela

                        List<MetadadoValorTabela> linhas = d.RevisaoAtual.ValoresTabelas.Where(mvt => mvt.ID_Metadado == metadado.Metadado.ID_Metadado).ToList();

                        JArray linhasTabela = new JArray();

                        foreach (MetadadoValorTabela mvt in linhas)
                        {
                            JObject objLinhaTabela = new JObject();

                            //foreach (ValorMetadado valor in mvt.ValoresMetadados)
                            foreach (var coluna in metadado.Metadado.ColunasMetadado)
                            {
                                string vlrCampo = "";

                                var valor = mvt.ValoresMetadados.FirstOrDefault(x => x.ID_Metadado == coluna.ID_MetadadoColuna);
                                if (valor != null)
                                {
                                    switch (valor.Metadado.Tipo)
                                    {
                                        case "string":
                                        case "memo":
                                            vlrCampo = valor.String;
                                            break;
                                        case "multi-valor":
                                            vlrCampo = _documentoService.RetornarValorMetadadoDocumento(valor);
                                            break;
                                        case "data":
                                        case "data_hora":
                                            vlrCampo = valor.Data.ToString();
                                            break;
                                        case "inteiro":
                                        case "numero":
                                            vlrCampo = valor.Inteiro.ToString();
                                            break;
                                        case "moeda":
                                            vlrCampo = valor.Float.ToString();
                                            break;
                                        case "lista":
                                            vlrCampo = _documentoService.RetornarValorMetadadoDocumento(valor);
                                            break;
                                        case "lista_sigla":
                                            vlrCampo = _documentoService.RetornarValorMetadadoDocumento(valor);
                                            break;
                                        case "checkbox":
                                            vlrCampo = valor.Bool.ToString();
                                            break;
                                        case "local_Estrutura":
                                            Pasta pasta = _pastaService.RetornarPasta(Convert.ToInt64(valor.Inteiro));
                                            if (pasta != null)
                                            {
                                                if (pasta.Sigla != "")
                                                {
                                                    vlrCampo = pasta.Sigla + "-" + pasta.Nome;
                                                }
                                                else
                                                {
                                                    vlrCampo = pasta.Nome;
                                                }
                                            }
                                            else
                                            {
                                                vlrCampo = "";
                                            }
                                            break;
                                    }
                                }

                                objLinhaTabela[coluna.Coluna.NomeInterno] = vlrCampo;
                            }

                            linhasTabela.Add(objLinhaTabela);
                        }


                        objeto[metadado.Metadado.NomeInterno] = linhasTabela;

                        #endregion
                    }
                }
            }

            #region Permissao alterar arquivo
            bool PodeAtualizarArquivoATV = false;
            bool PodeAtualizarArquivoUSU = false;
            if (usuario != null)
            {
                if (d.BPMInstancias.FirstOrDefault(i => i.Concluido == false) != null)
                {
                    PodeAtualizarArquivoATV = _arquivoService.VerificaPermissaoParaAtualizarArquivo(d, d.BPMInstancias.FirstOrDefault(i => i.Concluido == false));
                    PodeAtualizarArquivoUSU = _documentoService.VerificaResponsabilidade(d, usuario, d.BPMInstancias.FirstOrDefault(i => i.Concluido == false).ID_BPMInstancia);
                }

                if (PodeAtualizarArquivoATV && PodeAtualizarArquivoUSU)
                {
                    objeto["Permite_Editar_Arquivo"] = true;
                }
                else
                {
                    objeto["Permite_Editar_Arquivo"] = false;
                }

            }
            #endregion

            #region Arquivos do documento
            JArray objetoArquivos = new JArray();
            foreach (var arquivo in d.RevisaoAtual.Arquivos)
            {
                JObject objArquivo = new JObject();
                objArquivo["ID_Arquivo"] = arquivo.ID_Arquivo;
                objArquivo["Nome_Arquivo"] = arquivo.Nome + arquivo.Extensao;
                objetoArquivos.Add(objArquivo);
            }

            objeto["Arquivos"] = objetoArquivos;
            #endregion

            return objeto;
        }

        public async Task<JObject> Post()
        {
            var httpRequest = HttpContext.Current.Request;

            string NomeDocumento = "";
            string TipoItem = "";
            string Usuario = "";
            string Campos = "";
            string ID_Processo = "";
            string AtividadeInicial = "";
            string Comentarios = "";
            string Referencias = "";

            try
            {
                NomeDocumento = HttpContext.Current.Request.Params["NomeDocumento"];
                TipoItem = HttpContext.Current.Request.Params["TipoItem"];
                Usuario = HttpContext.Current.Request.Params["Usuario"];
                Campos = HttpContext.Current.Request.Params["Campos"];
                ID_Processo = HttpContext.Current.Request.Params["ID_Processo"]; // Id do processo de criacao, caso não for informado será criado no primeiro
                AtividadeInicial = HttpContext.Current.Request.Params["AtividadeInicial"]; //Nome interno da atividade que o doc deve ficar
                Comentarios = HttpContext.Current.Request.Params["Comentarios"]; //Comentarios do documento Pai
                Referencias = HttpContext.Current.Request.Params["Referencias"];
            }
            catch (Exception ex)
            {
                _servidorService.CriarLogErro(ex.Message, "Parâmetros - ItemsController - POST", "", 0, ex.StackTrace, "", "", "");
                var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                {
                    Content = new StringContent("Incorrect Parameters")
                };
                throw new HttpResponseException(resp);
            }

            long idTipoItem = 0;
            Int64.TryParse(TipoItem, out idTipoItem);

            if (idTipoItem == 0)
            {
                var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                {
                    Content = new StringContent("Incorrect Parameter: TipoItem")
                };
                throw new HttpResponseException(resp);
            }

            Documento documento = new Documento();
            string guid_acao = "";
            List<Email> EmailsNotificacao = new List<Email>();

            using (DbContextTransaction tran = context.Database.BeginTransaction(IsolationLevel.ReadCommitted))
            {
                try
                {
                    guid_acao = tran.GetHashCode().ToString();
                    #region Cria Item

                    TipoItem tipoItem = _projetoService.RetornarTipoItem(Convert.ToInt64(TipoItem));

                    if (string.IsNullOrEmpty(Campos))
                    {
                        Campos = "{}";
                    }

                    Dictionary<string, string> metadados = JsonConvert.DeserializeObject<Dictionary<string, string>>(Campos);

                    Projeto projeto = tipoItem.Projeto;
                    _documentoService.projetoSelecionado = projeto;

                    Usuario _usuario = _usuarioService.RetornarUsuario(Usuario);

                    if (_usuario == null)
                    {
                        _usuario = _usuarioService.RetornarUsuario(Thread.CurrentPrincipal.Identity.Name);
                    }

                    _documentoService.DefineUsuarioLogado(_usuario);
                    _projetoService.projetoSelecionado = projeto;
                    _bpmService.DefineUsuarioLogado(_usuario);
                    _bpmService.projetoSelecionado = projeto;

                    List<Metadado> listaMetadados = _projetoService.ListarMetadados(projeto, tipoItem.ID_TipoItem);

                    ListaValores valoresMetadados = new ListaValores();
                    //Guarda as linhas de tabelas
                    Dictionary<string, string> Tabelas = new Dictionary<string, string>();
                    Dictionary<long, List<long>> valoresFormRelacionado = new Dictionary<long, List<long>>();

                    Documento doc = new Documento();
                    doc.DataCriacao = DateTime.Now;
                    doc.Excluido = false;
                    doc.ID_Projeto = tipoItem.ID_Projeto;
                    doc.ID_TipoItem = tipoItem.ID_TipoItem;
                    doc.TipoItem = tipoItem;
                    doc.Nome = NomeDocumento;

                    foreach (var campos in metadados)
                    {
                        long idMetadado = 0;
                        string nome_interno = campos.Key;
                        string valor = campos.Value;

                        valor = TrataTextoJSON(valor);

                        Metadado m = null;
                        if (Int64.TryParse(campos.Key.Substring(1), out idMetadado))
                        {
                            m = listaMetadados.FirstOrDefault(mm => mm.ID_Metadado == idMetadado);
                        }
                        else
                        {
                            m = listaMetadados.FirstOrDefault(mm => mm.NomeInterno.ToUpper() == nome_interno.ToUpper());
                        }

                        if (m != null)
                        {
                            if (m.Tipo == "lista")
                            {
                                valoresMetadados.AddValor(m, valor);
                            }
                            else if (m.Tipo == "lista_sigla")
                            {
                                string sigla = valor.Split('-')[0];

                                ListasMetadado lm = _projetoService.RetornaItemListaMetadado(m, sigla);

                                if (lm != null)
                                    valor = lm.ID.ToString();
                                else
                                    valor = "0";

                                valoresMetadados.AddValor(m, valor);
                            }
                            else if (m.Tipo == "multi-valor")
                            {
                                List<string> Ids = new List<string>();

                                if (m.TipoLista == "usuarios")
                                {
                                    foreach (string v in valor.Split(','))
                                    {
                                        Usuario usu = _usuarioService.RetornarUsuario(v.TrimStart().TrimEnd());

                                        if(usu != null)
                                            Ids.Add(usu.ID_Usuario.ToString());
                                    }
                                }
                                else if (m.TipoLista == "grupos")
                                {
                                    foreach (string v in valor.Split(','))
                                    {
                                        Grupo grupo = _contaService.RetornarGrupoPeloNome(v.TrimStart().TrimEnd());

                                        if(grupo != null)
                                            Ids.Add(grupo.ID_Grupo.ToString());
                                    }
                                }
                                else {
                                    foreach (string v in valor.Split(','))
                                    {
                                        if (m.ListasMetadados.Count(lm => lm.Valor == v.TrimStart().TrimEnd()) > 0)
                                            Ids.Add(m.ListasMetadados.FirstOrDefault(lm => lm.Valor == v.TrimStart().TrimEnd()).ID.ToString());
                                    }
                                }

                                valoresMetadados.AddValor(m, String.Join(",", Ids.Distinct().ToArray()));
                            }
                            else if (m.Tipo == "local_Estrutura")
                            {
                                string sigla = valor.Split('-')[0];

                                Pasta pasta = _pastaService.RetornarPasta(m.Projeto, sigla, false);
                                if (pasta != null)
                                {
                                    valoresMetadados.AddValor(m, pasta.ID_Pasta.ToString());
                                }
                            }
                            else if (m.Tipo == "tabela")
                            {
                                Tabelas.Add(nome_interno, valor);
                            }
                            else if (m.Tipo == "form_relacionado")
                            {
                                if (!valoresFormRelacionado.ContainsKey(m.ID_Metadado))
                                {
                                    List<long> idsDocsReferenciados = new List<long>();
                                    MetadadoTiposItensRelacionados metTipoItemRel = m.TiposItensRelacionados.FirstOrDefault();

                                    if (metTipoItemRel != null)
                                    {
                                        if (!string.IsNullOrEmpty(valor))
                                        {
                                            foreach (string NomeDoc in valor.Split(','))
                                            {
                                                Documento docRef = _documentoService.RetornarDocumento(NomeDoc, metTipoItemRel.TipoItem.ID_Projeto, metTipoItemRel.TipoItem.ID_TipoItem);

                                                if (docRef != null)
                                                    idsDocsReferenciados.Add(docRef.ID_Documento);
                                            }
                                            valoresFormRelacionado.Add(m.ID_Metadado, idsDocsReferenciados);
                                        }
                                    }

                                }
                            }
                            else
                                valoresMetadados.AddValor(m, valor);
                        }
                    }

                    long IdProcesso = 0;
                    Processo processo = new Processo();
                    if (ID_Processo != "")
                    {
                        processo = projeto.Processos.FirstOrDefault(x => x.ID_Processo == Convert.ToInt64(ID_Processo));

                        if (processo != null)
                        {
                            IdProcesso = processo.ID_Processo;
                        }
                    }

                    bool gerarCodigo = false;
                    if (string.IsNullOrEmpty(NomeDocumento))
                    {
                        gerarCodigo = true;
                        doc.Nome = _projetoService.RetornaPrefixoNome(projeto, 0, valoresMetadados, (long)doc.ID_TipoItem);
                    }

                    documento = _documentoService.CriarDocumento(doc, valoresMetadados, _usuario, gerarCodigo, new List<Pasta>(), 0, true, IdProcesso);

                    RegistraInformacao(documento, _usuario, 1, httpRequest);

                    DocumentoMetadados DocCriadoPeloNome = _documentoService.ExtrairMetadadosNomeArquivo(documento.Nome, projeto, new List<Pasta>(), _usuario, documento.TipoItem, true, "", "", null, true, null, documento);
                    documento = DocCriadoPeloNome.Documento;

                    BPMInstancia instancia = doc.BPMInstancias.FirstOrDefault(c=> !c.Concluido);
                    BPMEvento evento = instancia.BPMElemento as BPMEvento;
                    if (evento != null)
                    {
                        BPMAcao acao = evento.BPMAcoes.FirstOrDefault();
                        #region  Cria um GUID da ação

                        string GUID_acao = tran.GetHashCode().ToString();

                        #endregion

                        _bpmService.ExecutarAcao(instancia.ID_BPMInstancia, acao.ID_Acao, _usuario.ID_Usuario, GUID_acao);
                    }
                    BPMInstancia ins = documento.BPMInstancias.FirstOrDefault(i => i.Concluido == false);
                    if (AtividadeInicial != "" && ID_Processo != "") // coloca documento na atividade escolhida
                    {
                        if (processo != null)
                        {
                            BPMElemento atividade = processo.BPMElementos.FirstOrDefault(a => a.NomeInterno == AtividadeInicial);
                            if (atividade != null)
                            {
                                ins.ID_BPMElemento = atividade.ID_BPMElemento;
                            }
                        }
                    }

                    EmailsNotificacao = _bpmService.BPMNovoEnvioEmail(_usuario, instancia, ins , guid_acao);
                    _documentoService.Save();

                    #region adiciona os valores as tabelas
                    foreach (var t in Tabelas)
                    {
                        string tabela = t.Key;
                        string linhas = t.Value;

                        Metadado metadadoTabela = listaMetadados.FirstOrDefault(mm => mm.NomeInterno.ToUpper() == tabela.ToUpper());

                        JArray jarr = JArray.Parse("[" + linhas + "]");
                        foreach (JObject content in jarr.Children<JObject>())
                        {
                            Dictionary<long, string> valores = new Dictionary<long, string>();

                            List<JProperty> campos = content.Properties().ToList();

                            foreach (JProperty prop in campos)
                            {
                                string nomeMetadadoLinha = prop.Name.ToString();
                                string valorLinha = prop.Value.ToString();

                                Metadado metadadoLinha = listaMetadados.FirstOrDefault(mm => mm.NomeInterno.ToUpper() == nomeMetadadoLinha.ToUpper());

                                if (metadadoLinha != null)
                                {
                                    if (metadadoLinha.Tipo == "lista")
                                    {
                                        valores.Add(metadadoLinha.ID_Metadado, valorLinha);
                                    }
                                    else if (metadadoLinha.Tipo == "lista_sigla")
                                    {
                                        string sigla = valorLinha.Split('-')[0];

                                        ListasMetadado lm = _projetoService.RetornaItemListaMetadado(metadadoLinha, sigla);
                                        if (lm != null)
                                            valorLinha = lm.ID.ToString();
                                        else
                                            valorLinha = "0";

                                        valores.Add(metadadoLinha.ID_Metadado, valorLinha);
                                    }
                                    else if (metadadoLinha.Tipo == "local_Estrutura")
                                    {
                                        string sigla = valorLinha.Split('-')[0];

                                        Pasta pasta = _pastaService.RetornarPasta(metadadoLinha.Projeto, sigla, false);
                                        if (pasta != null)
                                        {
                                            valores.Add(metadadoLinha.ID_Metadado, pasta.ID_Pasta.ToString());
                                        }
                                    }
                                    else
                                        valores.Add(metadadoLinha.ID_Metadado, valorLinha);

                                }
                            }

                            if (valores.Count > 0 && metadadoTabela != null)
                            {
                                MetadadoValorTabela mvt = _bpmService.adicionaLinhaTabela(metadadoTabela, documento, valores, _usuario);

                                GerenciarReservaDeSequencial(documento, projeto, mvt);
                            }
                        }

                    }
                    #endregion
                    #endregion

                    #region UploadArquivos
                    
                    string pathTemp = Path.GetTempPath();

                    var provider = new MultipartFormDataStreamProvider(pathTemp);

                    try
                    {
                        await Request.Content.ReadAsMultipartAsync(provider);
                    }
                    catch { }

                    if (provider.FileData.Count > 0)
                    {
                        var docfiles = new List<string>();

                        #region Limpa a revisão selecionada e adiciona os novos arquivos
                        Revisao revSel = null;
                        revSel = documento.RevisaoAtual;

                        var arquivos = revSel.Arquivos.ToList();

                        foreach (Arquivo arquivoAntigo in arquivos)
                        {
                            _arquivoService.RemoverArquivo(documento.ID_Documento, revSel.ID_Revisao, arquivoAntigo.ID_Arquivo);
                        }
                        #endregion

                        foreach (MultipartFileData postedFile in provider.FileData)
                        {
                            if (!string.IsNullOrEmpty(postedFile.Headers.ContentDisposition.FileName))
                            {
                                string fileName = postedFile.Headers.ContentDisposition.FileName;
                                if (fileName.StartsWith("\"") && fileName.EndsWith("\""))
                                {
                                    fileName = fileName.Trim('"');
                                }

                                if (fileName.Contains(@"/") || fileName.Contains(@"\"))
                                {
                                    fileName = Path.GetFileName(fileName);
                                }

                                if (!fileName.StartsWith("ComentNumero"))
                                {
                                    _arquivoService.CriarArquivoComConteudoRevisao(postedFile.LocalFileName, documento.RevisaoAtual, fileName, _usuario.ID_Usuario, 0);
                                    System.IO.File.Delete(postedFile.LocalFileName);
                                }
                            }
                        }
                    }
                    else if (!tipoItem.ArquivoOpcional)
                    {
                        _servidorService.CriarLogErro("Este tipo de item não permite criação de documentos sem arquivos.", "Erro - ItemsController - POST", "", 0, "", "", "", "");
                    }





                    #endregion

                    //Comentarios
                    CriaComentarios(documento, Comentarios, provider);

                    //referencias
                    CriaReferencias(projeto, Referencias);

                    #region passaFluxoAtividadeServico
                    foreach (BPMInstancia bpmInstancia in doc.BPMInstancias.Where(b => b.Concluido == false).ToList())
                    {
                        BPMElemento bpmElemento = bpmInstancia.BPMElemento;
                        if (bpmElemento is BPMAtividade)
                        {
                            BPMAtividade atividade = bpmElemento as BPMAtividade;

                            if (atividade.TipoAtividade == 4) //Serviço
                            {
                                if (atividade.GerarPDFArquivosEditaveis)
                                {
                                    _arquivoService.GerarPDFArquivos(doc, _usuario, atividade.GerarWatermark && !string.IsNullOrEmpty(doc.TipoItem.DescricaoWatermarkPDF), atividade.GerarCarimbo, false, null, atividade.GerarComentarios, atividade.GerarComentariosRevAnterior, atividade.ExcluirArquivosEditaveisRevisao, atividade.GerarComentariosPublicos);
                                }

                                foreach (Script script in atividade.Scripts)
                                {
                                    _scriptsService.ExecutaFuncaoScript(projeto, doc, _usuario, ScriptsEngine.EventosScriptItem.CustomFunction, script.Funcao);
                                }

                                _bpmService.Save();
                                _bpmService.projetoSelecionado = projeto;
                                _bpmService.IniciaExecutarAcao(bpmInstancia, atividade.BPMAcoes.FirstOrDefault(), _usuario);
                            }
                        }
                    }
                    #endregion


                    foreach (var item in valoresFormRelacionado)
                    {
                        foreach (long idDocRef in item.Value)
                        {
                            _documentoService.ReferenciaDocumento(documento.ID_Documento, idDocRef, item.Key);
                        }
                    }

                    tran.Commit();

                    JObject objeto = new JObject();

                    objeto = montaJSON(documento, 1, _usuario);

                    try
                    {
                        if (EmailsNotificacao.Count > 0 && projeto.ID_Projeto > 0)
                        {
                            _bpmService.projetoSelecionado = projeto;
                            _bpmService.EnviaEmailsPendencia(guid_acao);
                        }
                    }
                    catch (Exception ex)
                    {
                        _servidorService.CriarLogErro(ex.Message, "Erro - ItemsController - PUT - Envio e-mail", "", 0, ex.StackTrace, "", "", "");
                    }

                    return objeto;
                }
                catch (Exception ex)
                {
                    tran.Rollback();

                    _servidorService.CriarLogErro(ex.Message, "Erro - ItemsController - POST", "", 0, ex.StackTrace, "", "", "");

                    var resp = new HttpResponseMessage(HttpStatusCode.InternalServerError)
                    {
                        Content = new StringContent(ex.Message)
                    };
                    throw new HttpResponseException(resp);
                }
            }
        }

        private void GerenciarReservaDeSequencial(Documento documento, Projeto projeto, MetadadoValorTabela linhaTabela)
        {
            if ((int)documento.TipoItem.Natureza == (int)NaturezaTiposItem.Formulario && documento.TipoItem.GerenciarReservaDeSequencial)
            {
                TipoItem tipoItem = documento.TipoItem;

                SequencialService _sequencialService = new SequencialService(_unitOfWork);

                long idMetChave = (long)tipoItem.ID_MetadadoChaveReservaSequencial;
                long idMetQtd = (long)tipoItem.ID_MetadadoQuantidadeReservaSequencial;
                long idMetInicio = (long)tipoItem.ID_MetadadoInicioReservaSequencial;
                long idMetFim = (long)tipoItem.ID_MetadadoFimReservaSequencial;

                if (linhaTabela.ValoresMetadados.Count(vm => vm.ID_Metadado == idMetQtd) > 0)
                {
                    string chave = "";
                    int qtd = 0;
                    int inicio = 0;
                    int fim = 0;

                    if (documento.RevisaoAtual.ValoresMetadados.FirstOrDefault(vm => vm.ID_Metadado == idMetChave) != null)
                        chave = documento.RevisaoAtual.ValoresMetadados.FirstOrDefault(vm => vm.ID_Metadado == idMetChave).String;

                    if (linhaTabela.ValoresMetadados.FirstOrDefault(vm => vm.ID_Metadado == idMetQtd) != null)
                        qtd = (int)linhaTabela.ValoresMetadados.FirstOrDefault(vm => vm.ID_Metadado == idMetQtd).Inteiro;

                    if (linhaTabela.ValoresMetadados.FirstOrDefault(vm => vm.ID_Metadado == idMetInicio) != null)
                        inicio = (int)linhaTabela.ValoresMetadados.FirstOrDefault(vm => vm.ID_Metadado == idMetInicio).Inteiro;

                    if (linhaTabela.ValoresMetadados.FirstOrDefault(vm => vm.ID_Metadado == idMetFim) != null)
                        fim = (int)linhaTabela.ValoresMetadados.FirstOrDefault(vm => vm.ID_Metadado == idMetFim).Inteiro;

                    ValorMetadado vmInicio = linhaTabela.ValoresMetadados.FirstOrDefault(vm => vm.ID_Metadado == idMetInicio);
                    if (vmInicio != null)
                        vmInicio.Inteiro = inicio;
                    else
                    {
                        vmInicio = new ValorMetadado();
                        vmInicio.ID_Metadado = idMetInicio;
                        vmInicio.ID_Revisao = documento.RevisaoAtual.ID_Revisao;
                        vmInicio.Inteiro = inicio;
                        linhaTabela.ValoresMetadados.Add(vmInicio);
                    }

                    ValorMetadado vmFim = linhaTabela.ValoresMetadados.FirstOrDefault(vm => vm.ID_Metadado == idMetFim);
                    if (vmFim != null)
                        vmFim.Inteiro = fim;
                    else
                    {
                        ValorMetadado vmfim = new ValorMetadado();
                        vmfim.ID_Metadado = idMetFim;
                        vmfim.ID_Revisao = documento.RevisaoAtual.ID_Revisao;
                        vmfim.Inteiro = fim;
                        linhaTabela.ValoresMetadados.Add(vmfim);
                    }

                    _documentoService.Save();

                    long idVlrMetadadoInicio = 0;
                    if (vmInicio != null)
                        idVlrMetadadoInicio = vmInicio.ID;

                    _sequencialService.CriarReservaSequencial(projeto.ID_Projeto, (long)idVlrMetadadoInicio, chave, qtd, inicio, fim);
                }
            }
        }

        private void CriaReferencias(Projeto projeto, string Referencias)
        {
            #region Referencias

            if (Referencias != null)
            {
                if (Referencias.Length > 3)
                {
                    JArray JAC = JArray.Parse("[" + Referencias + "]");
                    foreach (JObject content in JAC.Children<JObject>())
                    {
                        List<JProperty> referencias = content.Properties().ToList();
                        string NomeDocumentoPai = "";
                        string TipoItemPai = "";

                        string TipoRef = "";
                        string UsuarioRef = "";

                        string NomeDocumentoFilho = "";
                        string TipoItemFilho = "";

                        foreach (JProperty prop in referencias)
                        {
                            if (prop.Name.ToString() == "NomeDocumentoPai")
                            {
                                NomeDocumentoPai = prop.Value.ToString();
                            }
                            else if (prop.Name.ToString() == "TipoItemPai")
                            {
                                TipoItemPai = prop.Value.ToString();
                            }
                            else if (prop.Name.ToString() == "TipoRef")
                            {
                                TipoRef = prop.Value.ToString();
                            }
                            else if (prop.Name.ToString() == "Usuario")
                            {
                                UsuarioRef = prop.Value.ToString();
                            }
                            else if (prop.Name.ToString() == "NomeDocumento")
                            {
                                NomeDocumentoFilho = prop.Value.ToString();
                            }
                            else
                            {
                                TipoItemFilho = prop.Value.ToString();
                            }
                        }

                        Usuario usuref = _usuarioService.RetornarUsuario(UsuarioRef);

                        TipoItem tipoItemFilho = _projetoService.RetornarTipoItem(Convert.ToInt64(TipoItemFilho));
                        TipoItem tipoItemPai = _projetoService.RetornarTipoItem(Convert.ToInt64(TipoItemPai));

                        projeto = tipoItemFilho.Projeto;
                        _documentoService.projetoSelecionado = projeto;
                        _documentoService.DefineUsuarioLogado(usuref);

                        Documento documentoFilho = _documentoService.RetornarDocumento(NomeDocumentoFilho, projeto.ID_Projeto, tipoItemFilho.ID_TipoItem, false);
                        Documento documentoPai = _documentoService.RetornarDocumento(NomeDocumentoPai, projeto.ID_Projeto, tipoItemPai.ID_TipoItem, false);

                        if (documentoFilho != null && documentoPai != null)
                        {
                            bool ret = _documentoService.ReferenciaDocumento(documentoPai.ID_Documento, documentoFilho.ID_Documento, 0, Convert.ToInt64(TipoRef));
                        }
                    }
                }
            }

            #endregion
        }

        private void CriaComentarios(Documento documento, string Comentarios, MultipartFormDataStreamProvider provider)
        {
            #region Comentarios

            JArray JAC = JArray.Parse("[" + Comentarios + "]");
            foreach (JObject content in JAC.Children<JObject>())
            {
                List<JProperty> comentarios = content.Properties().ToList();
                long numero = 0;
                string descricao = "";
                string loginUsuC = "";

                foreach (JProperty prop in comentarios)
                {
                    if (prop.Name.ToString() == "Numero")
                    {
                        numero = Convert.ToInt64(prop.Value.ToString());
                    }
                    else if (prop.Name.ToString() == "Descricao")
                    {
                        descricao = prop.Value.ToString();
                    }
                    else
                    {
                        loginUsuC = prop.Value.ToString();
                    }
                }

                Usuario usuComent = _usuarioService.RetornarUsuario(loginUsuC);

                Comentario coment = _comentarioService.CriarComentario(descricao, usuComent.ID_Usuario, (long)documento.ID_RevisaoAtual, true, null);

                #region UploadArquivos

                if (provider.FileData.Count > 0)
                {
                    var docfiles = new List<string>();

                    foreach (MultipartFileData postedFile in provider.FileData)
                    {
                        string fileName = postedFile.Headers.ContentDisposition.FileName;
                        if (fileName.StartsWith("\"") && fileName.EndsWith("\""))
                        {
                            fileName = fileName.Trim('"');
                        }

                        if (fileName.Contains(@"/") || fileName.Contains(@"\"))
                        {
                            fileName = Path.GetFileName(fileName);
                        }

                        if (fileName.StartsWith("ComentNumero" + numero + "."))
                        {
                            _arquivoService.CriarArquivoComConteudoRevisao(postedFile.LocalFileName, documento.RevisaoAtual, fileName, usuComent.ID_Usuario, 0);
                            System.IO.File.Delete(postedFile.LocalFileName);
                        }
                    }

                }

                #endregion
            }

            #endregion
        }

        // PUT: api/Itens/5
        public async Task<JObject> Put()
        {
            var httpRequest = HttpContext.Current.Request;

            string NomeDocumento = "";
            string TipoItem = "";
            string Usuario = "";
            string Campos = "";
            string RevisaoDocumento = "";
            string GerarRevisao = "";
            string CriarSeNaoExistir = "";
            string ID_Processo = "";

            string RevisaoOriginal = "";
            string NumeroSubRevisao = "";

            string Comentarios = "";
            string Referencias = "";
            string IDArquivo = "";

            string AtividadeDestino = "";

            string ID_Revisao = "";
            string ID_Documento = "";

            try
            {
                NomeDocumento = HttpContext.Current.Request.Params["NomeDocumento"];
                TipoItem = HttpContext.Current.Request.Params["TipoItem"];
                Usuario = HttpContext.Current.Request.Params["Usuario"];
                Campos = HttpContext.Current.Request.Params["Campos"];
                RevisaoDocumento = HttpContext.Current.Request.Params["RevisaoDocumento"];
                GerarRevisao = HttpContext.Current.Request.Params["GerarRevisao"];
                CriarSeNaoExistir = HttpContext.Current.Request.Params["CriarSeNaoExistir"];
                ID_Processo = HttpContext.Current.Request.Params["ID_Processo"]; // Id do processo de criacao, caso não for informado será criado no primeiro

                RevisaoOriginal = HttpContext.Current.Request.Params["RevisaoOriginal"];
                NumeroSubRevisao = HttpContext.Current.Request.Params["NumeroSubRevisao"];

                Comentarios = HttpContext.Current.Request.Params["Comentarios"]; //Comentarios do documento Pai
                Referencias = HttpContext.Current.Request.Params["Referencias"];
                IDArquivo = HttpContext.Current.Request.Params["ID_Arquivo"];

                AtividadeDestino = HttpContext.Current.Request.Params["AtividadeDestino"]; //Nome interno da atividade que o doc deve ficar

                ID_Revisao = HttpContext.Current.Request.Params["ID_Revisao"]; // Id da revisão, caso de uso do app mobile
                ID_Documento = HttpContext.Current.Request.Params["ID_Documento"];
            }
            catch (Exception ex)
            {
                _servidorService.CriarLogErro(ex.Message, "Parâmetros - ItemsController - PUT", "", 0, ex.StackTrace, "", "", "");

                var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                {
                    Content = new StringContent("Incorrect Parameters")
                };
                throw new HttpResponseException(resp);
            }

            //verifica se é importador
            if (httpRequest.UserAgent == "ImportadorGreendocs")
            {
                long idtipoitem = 0;
                long idprocesso = 0;
                if (Int64.TryParse(TipoItem, out idtipoitem) && Int64.TryParse(ID_Processo, out idprocesso))
                {
                    //pega o projeto pelo tipo de item
                    TipoItem tipoItem = _projetoService.RetornarTipoItem(idtipoitem);

                    //verifica se tem a constante
                    ConfigConstante constante = tipoItem.Projeto.Constantes.FirstOrDefault(c => c.Constante == "IMPORTADOR");
                    if (constante != null)
                    {
                        //verifica se o documento já existe
                        Documento documento = _documentoService.RetornarDocumento(NomeDocumento, tipoItem.ID_Projeto, tipoItem.ID_TipoItem, false);
                        if (documento == null)
                        {
                            return await CriaDocumentoSimplificado(NomeDocumento, Usuario, tipoItem, idprocesso, constante.String_Valor, Campos, RevisaoDocumento);
                        }
                    }
                }
            }

            List<Email> EmailsNotificacao = new List<Email>();
            Projeto projeto = new Projeto();
            JObject objeto = new JObject();
            string guid_acao = "";
  
            using (DbContextTransaction tran = context.Database.BeginTransaction(IsolationLevel.ReadCommitted))
            {
                string msgErro = "";
                guid_acao = tran.GetHashCode().ToString();

                try
                {
                    #region Alterar Item
                    TipoItem tipoItem = null;
                    Documento documento = null;
                    Revisao revisao = null;

                    string nome = NomeDocumento;

                    long idRevisao = 0;
                    long idDocumento = 0;

                    if (Int64.TryParse(ID_Documento, out idDocumento))
                    {
                        documento = _documentoService.RetornarDocumento(idDocumento);
                        revisao = documento.RevisaoAtual;
                        tipoItem = documento.TipoItem;
                        projeto = documento.Projeto;
                    }
                    else if (Int64.TryParse(ID_Revisao, out idRevisao))
                    {
                        revisao = _documentoService.RetornaRevisao(idRevisao);

                        if (revisao != null)
                        {
                            documento = revisao.Documento;
                            tipoItem = documento.TipoItem;
                            projeto = documento.Projeto;

                            if (tipoItem.RegraCodificacaoManual && tipoItem.Natureza == 2 && !string.IsNullOrEmpty(nome))
                            {
                                documento.Nome = nome;
                                revisao.NomeDocumento = nome;
                            }
                        }
                    }
                    else
                    {
                        tipoItem = _projetoService.RetornarTipoItem(Convert.ToInt64(TipoItem));
                        projeto = tipoItem.Projeto;
                        documento = _documentoService.RetornarDocumento(nome, projeto.ID_Projeto, tipoItem.ID_TipoItem, false);
                    }

                    _documentoService.projetoSelecionado = projeto;

                    Usuario _usuario = _usuarioService.RetornarUsuario(Usuario);

                    if (_usuario == null)
                    {
                        _usuario = _usuarioService.RetornarUsuario(Thread.CurrentPrincipal.Identity.Name);
                    }
                                        
                    _documentoService.DefineUsuarioLogado(_usuario);
                    _bpmService.DefineUsuarioLogado(_usuario);
                    _projetoService.projetoSelecionado = projeto;
                    _triggerService.DefineUsuarioLogado(_usuario);
                    _triggerService.projetoSelecionado = projeto;

                    if (string.IsNullOrEmpty(Campos))
                    {
                        Campos = "{}";
                    }

                    Dictionary<string, string> metadados = JsonConvert.DeserializeObject<Dictionary<string, string>>(Campos);

                    List<Metadado> listaMetadados = _projetoService.ListarMetadados(projeto, tipoItem.ID_TipoItem);

                    ListaValores valoresMetadados = new ListaValores();
                    List<Coluna> valoresBuiltIn = new List<Coluna>();
                    List<Pasta> locais = new List<Pasta>();
                    //Guarda as linhas de tabelas
                    Dictionary<string, string> Tabelas = new Dictionary<string, string>();
                    Dictionary<long, List<long>> valoresFormRelacionado = new Dictionary<long, List<long>>();

                    foreach (var campos in metadados)
                    {
                        msgErro = "";

                        long idMetadado = 0;
                        string nome_interno = campos.Key;
                        string valor = campos.Value;

                        valor = TrataTextoJSON(valor);
                        
                        Metadado m = null;
                        if (Int64.TryParse(campos.Key.Substring(1), out idMetadado))
                        {
                            m = listaMetadados.FirstOrDefault(mm => mm.ID_Metadado == idMetadado);
                        }
                        else
                        {
                            m = listaMetadados.FirstOrDefault(mm => mm.NomeInterno.ToUpper() == nome_interno.ToUpper());
                        }

                        if (m != null)
                        {
                            msgErro = string.Format("Metadado: {0} - Valor: {1};", m.Nome, valor);

                            #region Busca valor MetadadoPai na cascata

                            long idValorPaiCascata = 0;
                            if (m.Metadado1 != null)
                            {
                                //Busca o valorMetadado
                                ValorMetadado valorMetadadoPai = valoresMetadados.FirstOrDefault(vm => vm.ID_Metadado == m.Metadado1.ID_Metadado && vm.Inteiro != null);

                                if (valorMetadadoPai != null)
                                {
                                    idValorPaiCascata = (long)valorMetadadoPai.Inteiro;
                                }
                            }

                            #endregion
                            if (m.Tipo == "lista")
                            {
                                valoresMetadados.AddValor(m, valor, idValorPaiCascata, usuarios: _unitOfWork.Usuarios.GetAll());
                            }
                            else if (m.Tipo == "lista_sigla")
                            {
                                valoresMetadados.AddValor(m, valor, idValorPaiCascata);
                            }
                            else if (m.Tipo == "multi-valor")
                            {
                                List<string> Ids = new List<string>();

                                if (m.TipoLista == "usuarios")
                                {
                                    foreach (string v in valor.Split(','))
                                    {
                                        Usuario usu = _usuarioService.RetornarUsuario(v.TrimStart().TrimEnd());

                                        if (usu != null)
                                            Ids.Add(usu.ID_Usuario.ToString());
                                    }
                                }
                                else if (m.TipoLista == "grupos")
                                {
                                    foreach (string v in valor.Split(','))
                                    {
                                        Grupo grupo = _contaService.RetornarGrupoPeloNome(v.TrimStart().TrimEnd());

                                        if (grupo != null)
                                            Ids.Add(grupo.ID_Grupo.ToString());
                                    }
                                }
                                else
                                {
                                    foreach (string v in valor.Split(','))
                                    {
                                        if (m.ListasMetadados.Count(lm => lm.Valor == v.TrimStart().TrimEnd()) > 0)
                                            Ids.Add(m.ListasMetadados.FirstOrDefault(lm => lm.Valor == v.TrimStart().TrimEnd()).ID.ToString());
                                    }
                                }

                                valoresMetadados.AddValor(m, String.Join(",", Ids.Distinct().ToArray()));
                            }
                            else if (m.Tipo == "local_Estrutura")
                            {
                                string sigla = valor.Split('-')[0];

                                Pasta pasta = _pastaService.RetornarPasta(m.Projeto, sigla, false);
                                if (pasta != null)
                                {
                                    valoresMetadados.AddValor(m, pasta.ID_Pasta.ToString());
                                    locais.Add(pasta);
                                }
                                else {
                                    string nomePasta = valor;
                                    pasta = m.Projeto.Pastas.FirstOrDefault(p => p.Excluido == false && p.Nome.ToUpper() == nomePasta.ToUpper());

                                    if (pasta != null)
                                    {
                                        valoresMetadados.AddValor(m, pasta.ID_Pasta.ToString());
                                        locais.Add(pasta);
                                    }
                                }
                            }
                            else if (m.Tipo == "tabela")
                            {
                                Tabelas.Add(m.NomeInterno, valor);
                            }
                            else if (m.Tipo == "form_relacionado")
                            {
                                if (!valoresFormRelacionado.ContainsKey(m.ID_Metadado))
                                {
                                    List<long> idsDocsReferenciados = new List<long>();
                                    MetadadoTiposItensRelacionados metTipoItemRel = m.TiposItensRelacionados.FirstOrDefault();

                                    if (metTipoItemRel != null)
                                    {
                                        if (!string.IsNullOrEmpty(valor))
                                        {
                                            foreach (string NomeDoc in valor.Split(','))
                                            {
                                                Documento docRef = _documentoService.RetornarDocumento(NomeDoc, metTipoItemRel.TipoItem.ID_Projeto, metTipoItemRel.TipoItem.ID_TipoItem);

                                                if (docRef != null)
                                                    idsDocsReferenciados.Add(docRef.ID_Documento);
                                            }
                                            valoresFormRelacionado.Add(m.ID_Metadado, idsDocsReferenciados);
                                        }
                                    }
                                    
                                }
                            }
                            else
                                valoresMetadados.AddValor(m, valor);
                        }
                        else
                        {
                            if (!campos.Key.StartsWith("L")) //Chuncho só pra verificar se o que tá vindo é uma propriedade builtin mesmo (No caso um valor inteiro no nome_interno)
                            {
                                if (Int64.TryParse(campos.Key, out idMetadado))
                                {
                                    if (idMetadado > 0)
                                        valoresBuiltIn.Add(new Coluna { ID_Coluna = idMetadado, Descricao = valor });
                                }
                            }
                        }
                    }

                    msgErro = "";

                    bool documentoNovo = false;

                    #region Cria o documento caso não exista e tenha a opção true

                    if (documento == null && CriarSeNaoExistir != null)
                    {
                        documentoNovo = true;

                        documento = new Documento();
                        documento.DataCriacao = DateTime.Now;
                        documento.Excluido = false;
                        documento.ID_Projeto = tipoItem.ID_Projeto;
                        documento.ID_TipoItem = tipoItem.ID_TipoItem;
                        documento.TipoItem = tipoItem;
                        documento.Nome = NomeDocumento != null && NomeDocumento.Length > 255 ? NomeDocumento.Substring(0, 255) : NomeDocumento;

                        bool gerarCodigo = false;
                        if (string.IsNullOrEmpty(NomeDocumento))
                        {
                            gerarCodigo = true;
                            documento.Nome = _projetoService.RetornaPrefixoNome(projeto, 0, valoresMetadados, (long)documento.ID_TipoItem);
                        }

                        long IdProcesso = 0;
                        Processo processo = new Processo();
                        if (ID_Processo != "")
                        {
                            processo = projeto.Processos.FirstOrDefault(x => x.ID_Processo == Convert.ToInt64(ID_Processo));

                            if (processo != null)
                            {
                                IdProcesso = processo.ID_Processo;
                            }
                        }

                        //Cria no primeiro processo
                        documento = _documentoService.CriarDocumento(documento, valoresMetadados, _usuario, gerarCodigo, locais, 0, true, IdProcesso);                                             
                        string sequencialDocService = documento.Sequencial;                                           
                        DocumentoMetadados DocCriadoPeloNome = _documentoService.ExtrairMetadadosNomeArquivo(documento.Nome, projeto, locais, _usuario, documento.TipoItem, true, "", "", null, true, null, documento);
                        documento = DocCriadoPeloNome.Documento;                        
                        documento.Sequencial = sequencialDocService;

                        BPMInstancia instancia = documento.BPMInstancias.FirstOrDefault(c => !c.Concluido);
                        BPMEvento evento = instancia.BPMElemento as BPMEvento;

                        if (evento != null)
                        {
                            BPMAcao acao = evento.BPMAcoes.FirstOrDefault();
                            #region  Cria um GUID da ação

                            string GUID_acao = tran.GetHashCode().ToString();

                            #endregion

                            _bpmService.ExecutarAcao(instancia.ID_BPMInstancia, acao.ID_Acao, _usuario.ID_Usuario, GUID_acao);
                        }

                        _documentoService.Save();

                        RegistraInformacao(documento, _usuario, 1, httpRequest);
                    }
                    else if (documento == null && CriarSeNaoExistir == null)
                    {
                        tran.Rollback();
                        return objeto;
                    }

                    #endregion

                    if (projeto.ModoImportacao && !tipoItem.PermiteMultiArquivo)
                    {
                        GerarRevisao = "1";
                    }
                    else
                    {
                        if (revisao == null)
                        {
                            revisao = documento.Revisoes.FirstOrDefault(rev => rev.Numero == RevisaoDocumento);
                        }
                    }

                    if (revisao == null)
                    {
                        if (GerarRevisao != null && !documentoNovo) // gera nova revisão
                        {
                            _documentoService.CriarNovaRevisao(documento, _usuario, false, tipoItem.DescartarArquivosRevisaoAntiga);
                        }

                        documento.RevisaoAtual.Numero = RevisaoDocumento;
                        documento.RevisaoAtual.NumeroOriginal = RevisaoOriginal;
                        documento.RevisaoAtual.NumeroSubRevisao = NumeroSubRevisao;

                        revisao = documento.RevisaoAtual;
                    }

                    #region passaFluxoAtividadeServicoPUT
                    foreach (BPMInstancia bpmInstancia in documento.BPMInstancias.Where(b => b.Concluido == false).ToList())
                    {
                        BPMElemento bpmElemento = bpmInstancia.BPMElemento;
                        if (bpmElemento is BPMAtividade)
                        {
                            BPMAtividade atividade = bpmElemento as BPMAtividade;

                            if (atividade.TipoAtividade == 4 && documentoNovo) //Serviço
                            {
                                if (atividade.GerarPDFArquivosEditaveis)
                                {
                                    _arquivoService.GerarPDFArquivos(documento, _usuario, atividade.GerarWatermark && !string.IsNullOrEmpty(documento.TipoItem.DescricaoWatermarkPDF), atividade.GerarCarimbo, false, null, atividade.GerarComentarios, atividade.GerarComentariosRevAnterior, atividade.ExcluirArquivosEditaveisRevisao, atividade.GerarComentariosPublicos);
                                }

                                foreach (Script script in atividade.Scripts)
                                {
                                    _scriptsService.ExecutaFuncaoScript(projeto, documento, _usuario, ScriptsEngine.EventosScriptItem.CustomFunction, script.Funcao);
                                }

                                _bpmService.Save();
                                _bpmService.projetoSelecionado = projeto;
                                _bpmService.IniciaExecutarAcao(bpmInstancia, atividade.BPMAcoes.FirstOrDefault(), _usuario);
                            }
                        }
                    }

                    #endregion

                    _documentoService.Save();

                    if (documento != null)
                    {
                        _documentoService.AtualizarDocumento(documento, revisao.ValoresMetadados.ToList(), valoresMetadados, true, false);
                    }

                    if (AtividadeDestino != "") // coloca documento na atividade escolhida
                    {
                        BPMAtividade atividade = (BPMAtividade)documento.ProcessoOrigem.BPMElementos.FirstOrDefault(a => a.NomeInterno == AtividadeDestino);

                        if (atividade != null)
                        {
                            BPMInstancia instanciaAtual = documento.BPMInstancias.FirstOrDefault(i => i.Concluido == false);

                            _bpmService.ConcluirInstancia(instanciaAtual, null, _usuario);

                            List<BPMInstancia> instancias = _bpmService.AdicionarInstanciaAtividade(atividade, documento, _usuario);

                            documento.RevisaoAtual.ID_Situacao = atividade.ID_Situacao;

                            if (instancias.Count > 0)
                            {
                                BPMInstancia instancia = instancias.FirstOrDefault();
                                EmailsNotificacao = _bpmService.BPMNovoEnvioEmail(_usuario, instanciaAtual, instancia, guid_acao);
                            }

                        }
                    }

                    _documentoService.Save();

                    string pathTemp = Path.GetTempPath();

                    var provider = new MultipartFormDataStreamProvider(pathTemp);

                    #region UploadArquivos

                    if (HttpContext.Current.Request.Files.Count > 0 && documento != null)
                    {
                        await Request.Content.ReadAsMultipartAsync(provider);

                        var docfiles = new List<string>();

                        #region Limpa a revisão selecionada e adiciona os novos arquivos
                        Revisao revSel = null;
                        revSel = revisao;

                        var arquivos = revSel.Arquivos.ToList();

                        if (!string.IsNullOrEmpty(IDArquivo))
                        {
                            arquivos = arquivos.Where(a => a.ID_Arquivo == Convert.ToInt64(IDArquivo)).ToList();

                            foreach (Arquivo arquivoAntigo in arquivos)
                            {
                                _arquivoService.RemoverArquivo(documento.ID_Documento, revSel.ID_Revisao, arquivoAntigo.ID_Arquivo);
                            }
                        }
                        else
                        {
                            if (!tipoItem.PermiteMultiArquivo)
                            {
                                foreach (Arquivo arquivoAntigo in arquivos)
                                {
                                    _arquivoService.RemoverArquivo(documento.ID_Documento, revSel.ID_Revisao, arquivoAntigo.ID_Arquivo);
                                }
                            }
                        }
                        #endregion

                        foreach (MultipartFileData postedFile in provider.FileData)
                        {
                            if (!string.IsNullOrEmpty(postedFile.Headers.ContentDisposition.FileName))
                            {
                                string fileName = postedFile.Headers.ContentDisposition.FileName;
                                if (fileName.StartsWith("\"") && fileName.EndsWith("\""))
                                {
                                    fileName = fileName.Trim('"');
                                }
                                if (fileName.Contains(@"/") || fileName.Contains(@"\"))
                                {
                                    fileName = Path.GetFileName(fileName);
                                }

                                if (!fileName.StartsWith("ComentNumero") && !fileName.StartsWith("MetadadoArquivo"))
                                {
                                    _arquivoService.CriarArquivoComConteudoRevisao(postedFile.LocalFileName, revisao, fileName, _usuario.ID_Usuario, 0);
                                    System.IO.File.Delete(postedFile.LocalFileName);
                                }
                            }
                        }
                    }

                    #endregion

                    //campos arquivo
                    var camposArquivo = documento.RevisaoAtual.ValoresMetadados.Where(vm => vm.Metadado.Tipo == "arquivo" && vm.Inteiro.HasValue);
                    foreach (var campoArquivo in camposArquivo)
                    {
                        Arquivo arquivo = null;

                        #region UploadArquivos

                        if (HttpContext.Current.Request.Files.Count > 0 && documento != null)
                        {
                            var docfiles = new List<string>();

                            foreach (MultipartFileData postedFile in provider.FileData)
                            {
                                if (!string.IsNullOrEmpty(postedFile.Headers.ContentDisposition.FileName))
                                {
                                    string fileName = postedFile.Headers.ContentDisposition.FileName;
                                    if (fileName.StartsWith("\"") && fileName.EndsWith("\""))
                                    {
                                        fileName = fileName.Trim('"');
                                    }
                                    if (fileName.Contains(@"/") || fileName.Contains(@"\"))
                                    {
                                        fileName = Path.GetFileName(fileName);
                                    }

                                    if (fileName.StartsWith("MetadadoArquivo" + campoArquivo.Inteiro))
                                    {
                                        fileName = fileName.Replace("MetadadoArquivo" + campoArquivo.Inteiro + ".", "");
                                        arquivo = _arquivoService.CriarArquivoComConteudo(postedFile.LocalFileName, revisao, fileName, _usuario.ID_Usuario);
                                        System.IO.File.Delete(postedFile.LocalFileName);
                                    }
                                }
                            }
                        }

                        if (arquivo != null)
                        {
                            campoArquivo.Inteiro = arquivo.ID_Arquivo;
                        }

                        #endregion
                    }

                    #region adiciona os valores as tabelas


                    foreach (var t in Tabelas)
                    {
                        long idMetadado = 0;
                        string tabela = t.Key;
                        string linhas = t.Value;

                        Metadado metadadoTabela = null;
                        if (Int64.TryParse(t.Key, out idMetadado))
                        {
                            metadadoTabela = listaMetadados.FirstOrDefault(mm => mm.ID_Metadado == idMetadado);
                        }
                        else
                        {
                            metadadoTabela = listaMetadados.FirstOrDefault(mm => mm.NomeInterno.ToUpper() == tabela.ToUpper());
                        }

                        List<MetadadoValorTabela> mvts = revisao.ValoresTabelas.Where(v => v.Metadado.ID_Metadado == metadadoTabela.ID_Metadado).ToList();

                        foreach (MetadadoValorTabela mvt in mvts)
                        {
                            _documentoService.ExcluirValorMetadadoTabela(documento, mvt, false);
                        }

                        JArray jarr = JArray.Parse("[" + linhas + "]");
                        foreach (JObject content in jarr.Children<JObject>())
                        {
                            Dictionary<long, string> valores = new Dictionary<long, string>();

                            List<JProperty> campos = content.Properties().ToList();

                            foreach (JProperty prop in campos)
                            {
                                string nomeMetadadoLinha = prop.Name.ToString();
                                string valorLinha = prop.Value.ToString();

                                Metadado metadadoLinha = listaMetadados.FirstOrDefault(mm => mm.NomeInterno.ToUpper() == nomeMetadadoLinha.ToUpper());

                                if (metadadoLinha != null)
                                {
                                    if (metadadoLinha.Tipo == "arquivo")
                                    {
                                        Arquivo arquivo = null;

                                        //Em campos arquivos, se não estiver enviando um arquivo, o valorLinha será o id do arquivo já existente.
                                        //Se estiver enviando um arquivo, o valorLinha será um id único enviado também como parâmetro nome no próprio arquivo

                                        #region UploadArquivos

                                        if (HttpContext.Current.Request.Files.Count > 0 && documento != null)
                                        {
                                            var docfiles = new List<string>();

                                            foreach (MultipartFileData postedFile in provider.FileData)
                                            {
                                                if (!string.IsNullOrEmpty(postedFile.Headers.ContentDisposition.FileName))
                                                {
                                                    string fileName = postedFile.Headers.ContentDisposition.FileName;
                                                    if (fileName.StartsWith("\"") && fileName.EndsWith("\""))
                                                    {
                                                        fileName = fileName.Trim('"');
                                                    }
                                                    if (fileName.Contains(@"/") || fileName.Contains(@"\"))
                                                    {
                                                        fileName = Path.GetFileName(fileName);
                                                    }

                                                    //valorLinha será um id único
                                                    if (fileName.StartsWith("MetadadoArquivoTabela" + valorLinha))
                                                    {
                                                        fileName = fileName.Replace("MetadadoArquivoTabela" + valorLinha + ".", "");
                                                        arquivo = _arquivoService.CriarArquivoComConteudo(postedFile.LocalFileName, revisao, fileName, _usuario.ID_Usuario);
                                                        System.IO.File.Delete(postedFile.LocalFileName);
                                                    }
                                                }
                                            }
                                        }

                                        #endregion

                                        long idArquivo = 0;
                                        if (arquivo != null)
                                            idArquivo = arquivo.ID_Arquivo; //id do arquivo novo
                                        else
                                            Int64.TryParse(valorLinha, out idArquivo); //id do arquivo que já estava preenchido no campo

                                        valores.Add(metadadoLinha.ID_Metadado, idArquivo.ToString());
                                    }
                                    else if (metadadoLinha.Tipo == "lista" || (metadadoLinha.Tipo == "lista_sigla" && metadadoLinha.ExibeSigla == false))
                                    {
                                        if (metadadoLinha.TipoLista == "unidades" || metadadoLinha.TipoLista == "areas" || metadadoLinha.TipoLista == "funcoes")
                                        {
                                            ListaValores listaValores = new ListaValores();
                                            listaValores.AddValor(metadadoLinha, valorLinha);
                                            if (listaValores.Count > 0)
                                                valorLinha = listaValores.FirstOrDefault().Inteiro.ToString();
                                            else
                                                valorLinha = "0";
                                        }
                                        else
                                        {
                                            ListasMetadado lm = _projetoService.RetornaItemListaMetadadoPorValor(metadadoLinha, valorLinha);

                                            if (lm != null)
                                                valorLinha = lm.ID.ToString();
                                            else
                                                valorLinha = "0";
                                        }

                                        valores.Add(metadadoLinha.ID_Metadado, valorLinha);
                                    }
                                    else if (metadadoLinha.Tipo == "lista_sigla")
                                    {
                                        string sigla = valorLinha.Split('-')[0];

                                        ListasMetadado lm = _projetoService.RetornaItemListaMetadado(metadadoLinha, sigla);
                                        if (lm != null)
                                            valorLinha = lm.ID.ToString();
                                        else
                                            valorLinha = "0";

                                        valores.Add(metadadoLinha.ID_Metadado, valorLinha);
                                    }
                                    else if (metadadoLinha.Tipo == "local_Estrutura")
                                    {
                                        string sigla = valorLinha.Split('-')[0];

                                        Pasta pasta = _pastaService.RetornarPasta(metadadoLinha.Projeto, sigla, false);
                                        if (pasta != null)
                                        {
                                            valores.Add(metadadoLinha.ID_Metadado, pasta.ID_Pasta.ToString());
                                        }
                                    }
                                    else
                                        valores.Add(metadadoLinha.ID_Metadado, valorLinha);

                                }
                            }

                            if (valores.Count > 0 && metadadoTabela != null)
                            {
                                MetadadoValorTabela mvt = _bpmService.adicionaLinhaTabela(metadadoTabela, documento, valores, _usuario);

                                GerenciarReservaDeSequencial(documento, projeto, mvt);
                            }
                        }

                    }
                    #endregion

                    #endregion

                    //Comentarios
                    CriaComentarios(documento, Comentarios, provider);

                    //referencias
                    CriaReferencias(projeto, Referencias);

                    if (documento != null)
                    {
                        //Propriedades Builtin
                        foreach (var valorBuiltIn in valoresBuiltIn)
                        {
                            switch (valorBuiltIn.ID_Coluna)
                            {
                                case 2:
                                    Pasta pasta = _pastaService.RetornarPastaPelaDescricao(projeto, valorBuiltIn.Descricao);

                                    if (pasta != null)
                                    {
                                        documento.Locais.Add(pasta);
                                    }
                                    break;
                                case 3:
                                    documento.Titulo = valorBuiltIn.Descricao;
                                    break;
                                case 4:
                                    documento.RevisaoAtual.Numero = valorBuiltIn.Descricao;
                                    break;
                                case 5:
                                    DateTime dataAtualizacao = new DateTime();
                                    if (!string.IsNullOrEmpty(valorBuiltIn.Descricao) && DateTime.TryParse(valorBuiltIn.Descricao, out dataAtualizacao))
                                    {
                                        documento.RevisaoAtual.DataAtualizacao = dataAtualizacao;
                                    }
                                    break;
                                case 6:
                                    Usuario usuarioCriador = _usuarioService.RetornarUsuarioPeloNome(valorBuiltIn.Descricao);

                                    if (usuarioCriador != null)
                                    {
                                        documento.CriadoPor = usuarioCriador.ID_Usuario;
                                    }
                                    break;
                                case 7:
                                    Usuario usuarioAlteradoPor = _usuarioService.RetornarUsuarioPeloNome(valorBuiltIn.Descricao);

                                    if (usuarioAlteradoPor != null)
                                    {
                                        documento.RevisaoAtual.ModificadoPor = usuarioAlteradoPor.ID_Usuario;
                                    }
                                    break;
                                case 25:
                                    DateTime dataCriacao = new DateTime();
                                    if (!string.IsNullOrEmpty(valorBuiltIn.Descricao) && DateTime.TryParse(valorBuiltIn.Descricao, out dataCriacao))
                                    {
                                        documento.DataCriacao = dataCriacao;
                                    }
                                    break;
                                case 26:
                                    DateTime dataRevisao = new DateTime();
                                    if (!string.IsNullOrEmpty(valorBuiltIn.Descricao) && DateTime.TryParse(valorBuiltIn.Descricao, out dataRevisao))
                                    {
                                        revisao.Data = dataRevisao;
                                    }
                                    break;
                            }
                        }

                        _documentoService.AtualizaIndice(documento);
                    }

                    //metadados referencia (suporta apenas um item por campo)
                    foreach (var item in valoresFormRelacionado)
                    {
                        if (item.Value.Any())
                        {
                            string valorAtual = _documentoService.RetornarValorMetadadoDocumento(documento.ID_Documento, item.Key);

                            if (!string.IsNullOrEmpty(valorAtual))
                            {
                                ValorMetadado valorMetadado = documento.RevisaoAtual.ValoresMetadados.FirstOrDefault(vm => vm.ID_Metadado == item.Key);
                                if (valorMetadado != null)
                                {
                                    valorMetadado.String = "";
                                    _documentoService.Save();
                                }

                                string[] idsRefs = valorAtual.Split(',');
                                foreach (var id in idsRefs)
                                {
                                    long idRef = 0;
                                    Int64.TryParse(id, out idRef);
                                    if (idRef > 0)
                                    {
                                        _documentoService.apagarReferenciaIdRef(documento.ID_Documento, idRef, item.Key);
                                    }
                                }
                            }

                            foreach (long idDocRef in item.Value)
                            {
                                _documentoService.ReferenciaDocumento(documento.ID_Documento, idDocRef, item.Key);
                            }
                        }
                    }
                    
                    foreach (var refPai in documento.ReferenciasPais)
                    {
                        _documentoService.AtualizaIndice(refPai.Documentos, true);
                    }

                    objeto = montaJSON(documento, 1, _usuario);

                    tran.Commit();
                }
                catch (Exception ex)
                {
                    _documentoService.Save();

                    tran.Rollback();

                    _servidorService.CriarLogErro(ex.Message, "Erro - ItemsController - PUT", "", 0, ex.StackTrace, "", "", "");
                    
                    var resp = new HttpResponseMessage(HttpStatusCode.InternalServerError)
                    {
                        Content = new StringContent(ex.Message + " - " + msgErro),
                        ReasonPhrase = "There was an error creating the item"
                    };
                    throw new HttpResponseException(resp);
                }
            }

            try
            {
                if (EmailsNotificacao.Count > 0 && projeto.ID_Projeto > 0)
                {
                    _bpmService.projetoSelecionado = projeto;
                    _bpmService.EnviaEmailsPendencia(guid_acao);
                }
            }
            catch (Exception ex)
            {
                _servidorService.CriarLogErro(ex.Message, "Erro - ItemsController - PUT - Envio e-mail", "", 0, ex.StackTrace, "", "", "");
            }

            return objeto;
        }

        private async Task<JObject> CriaDocumentoSimplificado(string nomeDocumento, string loginUsuario, TipoItem tipoItem, long idProcesso, string nomeAtividade, string campos, string numeroRevisao)
        {
            JObject objeto = new JObject();
            string guid_acao = "";

            using (DbContextTransaction tran = context.Database.BeginTransaction(IsolationLevel.ReadCommitted))
            {
                string msgErro = "";
                guid_acao = tran.GetHashCode().ToString();

                try
                {
                    Usuario _usuario = _usuarioService.RetornarUsuario(loginUsuario);
                    
                    if (_usuario == null)
                    {
                        _usuario = _usuarioService.RetornarUsuario(Thread.CurrentPrincipal.Identity.Name);
                    }

                    _documentoService.DefineUsuarioLogado(_usuario);
                    //Cria o documento
                    Documento documento = new Documento();
                    documento.Nome = nomeDocumento;
                    documento.ID_Projeto = tipoItem.ID_Projeto;
                    documento.ID_TipoItem = tipoItem.ID_TipoItem;
                    documento.DataCriacao = DateTime.Now;
                    documento.CriadoPor = _usuario.ID_Usuario;
                    documento.ID_ProcessoOrigem = idProcesso;

                    _unitOfWork.Documentos.Insert(documento);

                    //Cria a revisão
                    Revisao revisao = new Revisao();
                    revisao.Data = DateTime.Now;
                    revisao.Documento = documento;
                    revisao.Usuario = _usuario;
                    revisao.RevisadoPor = _usuario.ID_Usuario;
                    revisao.NomeDocumento = nomeDocumento;
                    
                    if (!string.IsNullOrEmpty(numeroRevisao))
                        revisao.Numero = numeroRevisao;
                    else
                        revisao.Numero = tipoItem.Projeto.IniciaRevisaoEm.HasValue ? tipoItem.Projeto.IniciaRevisaoEm.ToString() : "";

                    _unitOfWork.Revisoes.Insert(revisao);
                    _unitOfWork.Save();

                    //Atribui revisão ao documento criado
                    documento.RevisaoAtual = revisao;

                    //coloca o documento na atividade
                    BPMInstancia instancia = new BPMInstancia();
                    BPMElemento bpmElemento = _unitOfWork.BPMElementos.Query(e => e.NomeInterno == nomeAtividade && e.ID_Processo == idProcesso).FirstOrDefault();
                    instancia.ID_BPMElemento = bpmElemento.ID_BPMElemento;
                    instancia.ID_Documento = documento.ID_Documento;
                    instancia.ID_Revisao = revisao.ID_Revisao;
                    instancia.DataInicio = DateTime.Now;

                    //definir responsável
                    instancia.Responsaveis.Add(_usuario);

                    _unitOfWork.BPMInstancias.Insert(instancia);

                    //registra evento criacao
                    RegistraInformacao(documento, _usuario, 1, HttpContext.Current.Request);

                    _unitOfWork.Save();

                    //Preenche metadados
                    if (string.IsNullOrEmpty(campos))
                    {
                        campos = "{}";
                    }

                    Dictionary<string, string> metadados = JsonConvert.DeserializeObject<Dictionary<string, string>>(campos);
                    List<Metadado> listaMetadados = _projetoService.ListarMetadados(tipoItem.Projeto, tipoItem.ID_TipoItem);
                    ListaValores valoresMetadados = new ListaValores();
                    List<Pasta> locais = new List<Pasta>();

                    foreach (var campo in metadados)
                    {
                        msgErro = "";

                        long idMetadado = 0;
                        string nome_interno = campo.Key;
                        string valor = campo.Value;

                        valor = TrataTextoJSON(valor);

                        Metadado m = null;
                        if (Int64.TryParse(campo.Key.Substring(1), out idMetadado))
                        {
                            m = listaMetadados.FirstOrDefault(mm => mm.ID_Metadado == idMetadado);
                        }
                        else
                        {
                            m = listaMetadados.FirstOrDefault(mm => mm.NomeInterno.ToUpper() == nome_interno.ToUpper());
                        }

                        if (m != null)
                        {
                            msgErro = string.Format("Metadado: {0} - Valor: {1};", m.Nome, valor);

                            #region Busca valor MetadadoPai na cascata

                            long idValorPaiCascata = 0;
                            if (m.Metadado1 != null)
                            {
                                //Busca o valorMetadado
                                ValorMetadado valorMetadadoPai = valoresMetadados.FirstOrDefault(vm => vm.ID_Metadado == m.Metadado1.ID_Metadado && vm.Inteiro != null);

                                if (valorMetadadoPai != null)
                                {
                                    idValorPaiCascata = (long)valorMetadadoPai.Inteiro;
                                }
                            }

                            #endregion
                            if (m.Tipo == "lista")
                            {
                                valoresMetadados.AddValor(m, valor, idValorPaiCascata, usuarios: _unitOfWork.Usuarios.GetAll());
                            }
                            else if (m.Tipo == "lista_sigla")
                            {
                                valoresMetadados.AddValor(m, valor, idValorPaiCascata);
                            }
                            else if (m.Tipo == "multi-valor")
                            {
                                List<string> Ids = new List<string>();

                                if (m.TipoLista == "usuarios")
                                {
                                    foreach (string v in valor.Split(','))
                                    {
                                        Usuario usu = _usuarioService.RetornarUsuario(v.TrimStart().TrimEnd());

                                        if (usu != null)
                                            Ids.Add(usu.ID_Usuario.ToString());
                                    }
                                }
                                else if (m.TipoLista == "grupos")
                                {
                                    foreach (string v in valor.Split(','))
                                    {
                                        Grupo grupo = _contaService.RetornarGrupoPeloNome(v.TrimStart().TrimEnd());

                                        if (grupo != null)
                                            Ids.Add(grupo.ID_Grupo.ToString());
                                    }
                                }
                                else
                                {
                                    foreach (string v in valor.Split(','))
                                    {
                                        if (m.ListasMetadados.Count(lm => lm.Valor == v.TrimStart().TrimEnd()) > 0)
                                            Ids.Add(m.ListasMetadados.FirstOrDefault(lm => lm.Valor == v.TrimStart().TrimEnd()).ID.ToString());
                                    }
                                }

                                valoresMetadados.AddValor(m, String.Join(",", Ids.Distinct().ToArray()));
                            }
                            else if (m.Tipo == "local_Estrutura")
                            {
                                string sigla = valor.Split('-')[0];

                                Pasta pasta = _pastaService.RetornarPasta(m.Projeto, sigla, false);
                                if (pasta != null)
                                {
                                    valoresMetadados.AddValor(m, pasta.ID_Pasta.ToString());
                                    locais.Add(pasta);
                                }
                                else
                                {
                                    string nomePasta = valor;
                                    pasta = m.Projeto.Pastas.FirstOrDefault(p => p.Excluido == false && p.Nome.ToUpper() == nomePasta.ToUpper());

                                    if (pasta != null)
                                    {
                                        valoresMetadados.AddValor(m, pasta.ID_Pasta.ToString());
                                        locais.Add(pasta);
                                    }
                                }
                            }
                            else if (m.Tipo == "form_relacionado")
                            {
                                foreach (String Doc in valor.Split(','))
                                {
                                    MetadadoTiposItensRelacionados metItemRelacionado = m.TiposItensRelacionados.FirstOrDefault();

                                    if (metItemRelacionado != null)
                                    {
                                        Documento docReferenciado = _documentoService.RetornarDocumento(Doc, documento.ID_Projeto, metItemRelacionado.TipoItem.ID_TipoItem);

                                        if (docReferenciado != null)
                                        {
                                            _documentoService.ReferenciaDocumento(documento, docReferenciado, m.ID_Metadado);

                                            Referencia refDocs = _documentoService.RetornaReferencia(documento.ID_Documento, docReferenciado.ID_Documento);

                                            ValorMetadado vm = new ValorMetadado { ID_Metadado = m.ID_Metadado, ID_Documento = documento.ID_Documento, ID_Revisao = (long)documento.ID_RevisaoAtual, Metadado = m };

                                            vm.ValoresMetadadosMultiValor.Add(new ValoresMetadadosMultiValor()
                                            {
                                                ID_Referencia = refDocs.ID_Referencia
                                            });

                                            _unitOfWork.Save();
                                        }
                                    }
                                }
                            }
                            else
                                valoresMetadados.AddValor(m, valor);
                        }
                    }

                    _documentoService.AtualizarDocumento(documento, new List<ValorMetadado>(), valoresMetadados, false);
                    
                    //upload do arquivo
                    string pathTemp = Path.GetTempPath();
                    var provider = new MultipartFormDataStreamProvider(pathTemp);

                    if (HttpContext.Current.Request.Files.Count > 0 && documento != null)
                    {
                        await Request.Content.ReadAsMultipartAsync(provider);

                        foreach (MultipartFileData postedFile in provider.FileData)
                        {
                            if (!string.IsNullOrEmpty(postedFile.Headers.ContentDisposition.FileName))
                            {
                                string fileName = postedFile.Headers.ContentDisposition.FileName;
                                if (fileName.StartsWith("\"") && fileName.EndsWith("\""))
                                {
                                    fileName = fileName.Trim('"');
                                }
                                if (fileName.Contains(@"/") || fileName.Contains(@"\"))
                                {
                                    fileName = Path.GetFileName(fileName);
                                }

                                if (!fileName.StartsWith("ComentNumero") && !fileName.StartsWith("MetadadoArquivo"))
                                {
                                    //_arquivoService.CriarArquivoComConteudoRevisao(postedFile.LocalFileName, revisao, fileName, _usuario.ID_Usuario, 0);
                                    Arquivo arquivo = _arquivoService.CriarArquivo(fileName, revisao.Documento.ID_Projeto, _usuario.ID_Usuario);

                                    if (_servidorService.RetornaTipoStorage() == TipoStorage.Azure) //Fiz este if pois carregar o arquivo em memória é desnecessário quando for filestorage
                                    {
                                        using (FileStream stream = File.OpenRead(postedFile.LocalFileName))
                                        {
                                            _arquivoService.AppendConteudoArquivo(arquivo, stream);
                                        }
                                    }
                                    else
                                    {
                                        System.IO.File.Move(postedFile.LocalFileName, arquivo.Path);
                                        _arquivoService.SalvarArquivo(arquivo, arquivo.Path);
                                    }

                                    revisao.Arquivos.Add(arquivo);
                                    _unitOfWork.Save();

                                    System.IO.File.Delete(postedFile.LocalFileName);
                                }
                            }
                        }
                    }

                    _documentoService.AtualizaIndice(documento);

                    objeto = montaJSON(documento, 1, _usuario);

                    tran.Commit();
                }
                catch (Exception ex)
                {
                    tran.Rollback();

                    _servidorService.CriarLogErro(ex.Message, "Erro - ItemsController - PUT", "", 0, ex.StackTrace, "", "", "");

                    var resp = new HttpResponseMessage(HttpStatusCode.InternalServerError)
                    {
                        Content = new StringContent(ex.Message + " - " + msgErro),
                        ReasonPhrase = "There was an error creating the item"
                    };
                    throw new HttpResponseException(resp);
                }
            }

            return objeto;
        }

        private string TrataTextoJSON(string texto) //Realiza um tratamento sobre os caracteres especiais, além de tratar [{,.:;/ entre ourtos
        {
            string textoLimpo = "";

            if (texto != null)
            {
                textoLimpo = texto;
            }

            textoLimpo = HttpUtility.UrlDecode(textoLimpo);

            return textoLimpo;
        }

        // DELETE: api/Itens/5
        public void Delete(int id)
        {
            throw new NotImplementedException();
        }

        private void RegistraInformacao(Documento documento, Usuario usuario, long evento, HttpRequest httpRequest)
        {
            Projeto projeto = documento.Projeto;
            
            try
            {
                Processo processo = documento.TipoItem.Processos.FirstOrDefault();

                string IP_Adress = httpRequest.ServerVariables["REMOTE_ADDR"];
                DateTime dataEvento = DateTime.Now;

                string obs = "";
                if (httpRequest.UserAgent == "GreendocsApp")
                    obs = "Criado via aplicativo";

                RegistroInformacoes registro = _contaService.CriarRegistro(documento, evento, obs, usuario, dataEvento, projeto, processo, IP_Adress);
            }
            catch (Exception ex)
            {
                _servidorService.CriarLogErro(ex.Message, "DocumentoController.cs", "RegistraInformacao", 0, ex.StackTrace, "", "", "", projeto.ID_Projeto, usuario.ID_Usuario);
            }
        }

        public JObject Excluir(long id)
        {
            JObject retorno = new JObject();
            using (DbContextTransaction tran = context.Database.BeginTransaction(IsolationLevel.ReadCommitted))
            {
                Usuario usuario = _usuarioService.RetornarUsuario(Thread.CurrentPrincipal.Identity.Name);
                Documento doc = _documentoService.RetornarDocumento(id);
                if (doc != null && usuario.Services)
                {
                    try
                    {
                        RegistraInformacao(doc, usuario, 6, HttpContext.Current.Request);
                        retorno["Log"] = "ID - {" + doc.ID_Documento + "} " + doc.Nome + " foi excluido com sucesso";
                        _projetoService.ExcluirItem(id, usuario);
                        tran.Commit();
                        retorno["Excluido"] = true;
                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            tran.Rollback();
                        }
                        catch (Exception)
                        {

                        }                      

                        using (UnitOfWork unitOfWork = new UnitOfWork())
                        {

                            ServidorService servidorService = new ServidorService(unitOfWork);
                            servidorService.CriarLogErro(ex.Message, "Erro - ItemsController - Excluir Documento", "Excluir", 0, ex.StackTrace, "", "", "");
                        }

                        retorno["Excluido"] = false;
                        retorno["Log"] = "Houve um erro na exclusão do documento ID - {"+doc.ID_Documento +"} " + doc.Nome;
                        retorno["Message"] = ex.Message;
                        retorno["StackTrace"] = ex.StackTrace;

                    }
                }
                else
                {
                    retorno["Log"] = "O usuário logado não possui permissão.";
                    retorno["Excluido"] = true;
                }
            
            }
            return retorno;
        }
    }

}