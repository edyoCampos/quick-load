using GreenDoc.Data;
using GreenDoc.Infrastructure;
using GreenDoc.Models;
using GreenDoc.Services;
using GreenDoc.Services.Promotor;
using GreenDoc.Web.ViewModels;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;

namespace GreenDoc.Web.Api.Controllers
{
    [MyBasicAuthenticationFilter]
    public class FilesController : ApiController
    {
        public UnitOfWork unitOfWork;
        public GreenDocContext context;
        public ContaService _contaService;
        public UsuarioService _usuarioService;
        public DocumentoService _documentoService;
        public ProjetoService _projetoService;
        public BPMService _bpmService;
        public ServidorService _servidorService;
        public ArquivoService _arquivoService;
        public TriggerService _triggerService;

        public FilesController()
        {
            context = new GreenDocContext();
            unitOfWork = new UnitOfWork(context);

            _contaService = new ContaService(unitOfWork);
            _usuarioService = new UsuarioService(unitOfWork);
            _triggerService = new TriggerService(unitOfWork);
            _documentoService = new DocumentoService(unitOfWork, _triggerService);
            _projetoService = new ProjetoService(unitOfWork);
            _bpmService = new BPMService(unitOfWork, _triggerService);
            _servidorService = new ServidorService(unitOfWork);
            _arquivoService = new ArquivoService(unitOfWork, _triggerService);
        }

        public HttpResponseMessage Get()
        {
            string NomeDocumento = HttpContext.Current.Request.Params["NomeDocumento"];
            string TipoItem = HttpContext.Current.Request.Params["TipoItem"];
            string ID_Arquivo = HttpContext.Current.Request.Params["ID_Arquivo"];
            string Ambiente = HttpContext.Current.Request.Params["Ambiente"];
            string ID_Documento = HttpContext.Current.Request.Params["ID_Documento"];
            string ID_Revisao = HttpContext.Current.Request.Params["ID_Revisao"];
            string gerarPDF = HttpContext.Current.Request.Params["PDF"];

            Documento documento = new Documento();
            Revisao revisao = new Revisao();
            List<Arquivo> arquivos = new List<Arquivo>();

            long id_tipoitem = 0;
            long id_projeto = 0;

            long idRev = 0;
            Int64.TryParse(ID_Revisao, out idRev);

            if (idRev > 0)
            {
                revisao = _documentoService.RetornaRevisao(idRev);
                if (revisao != null && !revisao.Documento.Excluido)
                {
                    arquivos = revisao.Arquivos.ToList();
                    documento = revisao.Documento;

                    if (documento.TipoItem.PermiteMultiArquivo || revisao.Arquivos.Count > 1)
                    {
                        if (!string.IsNullOrEmpty(ID_Arquivo))
                        {
                            long idArq = 0;
                            Guid guidFid;
                            if (Int64.TryParse(ID_Arquivo, out idArq))
                            {
                                arquivos = arquivos.Where(a => a.ID_Arquivo == idArq).ToList();
                            }
                            else if (Guid.TryParse(ID_Arquivo, out guidFid))
                            {
                                arquivos = arquivos.Where(a => a.FID == guidFid).ToList();
                            }
                        }
                    }                        
                }                
            }
            else if (!string.IsNullOrEmpty(ID_Documento))
            {
                long idDoc = 0;
                if (Int64.TryParse(ID_Documento, out idDoc))
                {
                    documento = _documentoService.RetornarDocumento(idDoc);
                    if (documento != null && !documento.Excluido)
                    {
                        arquivos = documento.RevisaoAtual.Arquivos.ToList();
                        revisao = documento.RevisaoAtual;
                    }
                }
            }
            else if ((!string.IsNullOrEmpty(TipoItem) || !string.IsNullOrEmpty(Ambiente)) && !string.IsNullOrEmpty(NomeDocumento))
            {
                if (!string.IsNullOrEmpty(Ambiente))
                {
                    if (!Int64.TryParse(Ambiente, out id_projeto))
                    {
                        var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                        {
                            Content = new StringContent("Parameter missing or in incorrect format: Ambiente")
                        };
                        throw new HttpResponseException(resp);
                    }

                    Projeto projeto = _projetoService.RetornarProjeto(id_projeto);

                    if (projeto == null)
                    {
                        var resp = new HttpResponseMessage(HttpStatusCode.NotFound)
                        {
                            Content = new StringContent(string.Format("No environment with ID = {0}", id_projeto)),
                            ReasonPhrase = "Environment Not Found"
                        };
                        throw new HttpResponseException(resp);
                    }

                    _documentoService.projetoSelecionado = projeto;

                    documento = _documentoService.RetornarDocumento(NomeDocumento, projeto.ID_Projeto);

                    if (documento == null)
                        documento = _documentoService.RetornarDocumento(NomeDocumento, projeto.ID_Projeto, true);

                    if(documento != null)
                    {
                        revisao = documento.RevisaoAtual;
                    }
                }
                else
                {
                    if (!Int64.TryParse(TipoItem, out id_tipoitem))
                    {
                        var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                        {
                            Content = new StringContent("Parameter missing or in incorrect format: TipoItem")
                        };
                        throw new HttpResponseException(resp);
                    }

                    TipoItem tipoItem = _projetoService.RetornarTipoItem(id_tipoitem);

                    if (tipoItem == null)
                    {
                        var resp = new HttpResponseMessage(HttpStatusCode.NotFound)
                        {
                            Content = new StringContent(string.Format("No item type with ID = {0}", id_tipoitem)),
                            ReasonPhrase = "Item Type Not Found"
                        };
                        throw new HttpResponseException(resp);
                    }

                    Projeto projeto = tipoItem.Projeto;

                    _documentoService.projetoSelecionado = projeto;

                    documento = _documentoService.RetornarDocumento(NomeDocumento, projeto.ID_Projeto, tipoItem.ID_TipoItem, false);

                    if (documento == null)
                        documento = _documentoService.RetornarDocumento(NomeDocumento, projeto.ID_Projeto, tipoItem.ID_TipoItem, true);

                    if (documento != null)
                    {
                        revisao = documento.RevisaoAtual;
                    }
                }

                if (!string.IsNullOrEmpty(ID_Arquivo))
                {
                    long idArq = 0;
                    Guid guidFid;
                    if (Int64.TryParse(ID_Arquivo, out idArq))
                        arquivos = arquivos.Where(a => a.ID_Arquivo == idArq).ToList();
                    else if (Guid.TryParse(ID_Arquivo, out guidFid))
                        arquivos = arquivos.Where(a => a.FID == guidFid).ToList();
                }
                else
                {
                    arquivos = documento.RevisaoAtual.Arquivos.ToList();
                }
            }
            else if (!string.IsNullOrEmpty(ID_Arquivo))
            {
                long idArq = 0;
                Guid guidFid;
                if (Int64.TryParse(ID_Arquivo, out idArq))
                {
                    Arquivo arquivo = _arquivoService.RetornarArquivo(idArq);
                    if (arquivo != null)
                        arquivos.Add(arquivo);
                }
                else if (Guid.TryParse(ID_Arquivo, out guidFid))
                {
                    Arquivo arquivo = _arquivoService.RetornarArquivo(guidFid);
                    if (arquivo != null)
                        arquivos.Add(arquivo);
                }
                else
                {
                    var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                    {
                        Content = new StringContent("Parameter missing or in incorrect format: ID_Arquivo")
                    };
                    throw new HttpResponseException(resp);
                }

                if (arquivos.Count > 0)
                {
                    if (arquivos.Count(x => x.Revisoes.Count > 0) > 0)
                    {
                        revisao = arquivos.First(x => x.Revisoes.Count > 0).Revisoes.First(); 
                        documento = revisao.Documento;
                    }
                    else
                    {                    
                        Arquivo arquivo = arquivos.FirstOrDefault();
                        ValorMetadado valorMetadado = unitOfWork.ValoresMetadados.Query(vm => vm.Inteiro == arquivo.ID_Arquivo && vm.Metadado.Tipo == "arquivo").FirstOrDefault();
                        if (valorMetadado != null)
                        {
                            revisao = valorMetadado.Revisao;
                            documento = revisao.Documento;
                        }
                    }
                }
            }

            HttpResponseMessage result = null;

            Usuario usuario = _usuarioService.RetornarUsuario(Thread.CurrentPrincipal.Identity.Name);

            BPMInstancia instancia = _bpmService.RetornaInstanciaResponsavel(usuario.ID_Usuario, documento.ID_Documento);

            bool podeBaixar = false;
            
            if(usuario.Services == true)
            {
                podeBaixar = true;
            }
            else
            {
                podeBaixar = _bpmService.VerificaAcessoInstancia(instancia, usuario);

                if (podeBaixar)
                    podeBaixar = _arquivoService.PodeBaixarArquivo(instancia, _documentoService.VerificaResponsabilidade(documento, usuario, instancia.ID_BPMInstancia), (long)instancia.ID_Revisao, usuario);
            }
            
            if (podeBaixar)
            {
                if (arquivos.Count == 1)
                {
                    var arq = arquivos.First();

                    byte[] conteudo = null;

                    bool pdf = false;
                    bool.TryParse(gerarPDF, out pdf);

                    if (arq.Extensao.ToLower() != ".pdf" && pdf)
                        conteudo = _arquivoService.GerarPDFArquivo(documento, arq, usuario, false, false, false, false).conteudoArquivo;
                    else
                        conteudo = _arquivoService.RetornarConteudoArquivo(arq);

                    result = Download(Request, conteudo, arq.Nome + arq.Extensao);
                }
                else if (arquivos.Count > 1)
                {
                    string numeroRevisao = "";
                    if (!string.IsNullOrEmpty(revisao.Numero))
                        numeroRevisao = revisao.Numero.Replace(" ", "_").Replace("/", "_");

                    string nomeArquivoZip = "Arquivos_" + documento.Nome.Replace(" ", "_").Replace("/", "_") + "_Rev" + numeroRevisao + "_" + revisao.ID_Revisao + ".zip";

                    using (var compressedFile = new MemoryStream())
                    {
                        using (var zipArchive = new ZipArchive(compressedFile, ZipArchiveMode.Update, false))
                        {
                            foreach (Arquivo arquivo in arquivos)
                            {
                                byte[] conteudoArquivo = _arquivoService.RetornarConteudoArquivo(arquivo);
                                if (conteudoArquivo != null && conteudoArquivo.Length > 0)
                                {
                                    var zipEntry = zipArchive.CreateEntry(arquivo.Nome + arquivo.Extensao);

                                    using (var originalFileStream = new MemoryStream(conteudoArquivo))
                                    {
                                        using (var zipEntryStream = zipEntry.Open())
                                        {
                                            originalFileStream.CopyTo(zipEntryStream);
                                        }
                                    }
                                }                                
                            }
                        }
                        result = Download(Request, compressedFile.ToArray(), nomeArquivoZip);
                    }
                }
            }
            else
            {
                var resp = new HttpResponseMessage(HttpStatusCode.Forbidden)
                {
                    Content = new StringContent("Usuário sem permissão para download")
                };
                throw new HttpResponseException(resp);
            }
            
            return result;
        }
        
        private static HttpResponseMessage Download(HttpRequestMessage Request, byte[] conteudoArquivo, string nomeArquivo)
        {
            HttpResponseMessage result = null;
            if (conteudoArquivo != null && conteudoArquivo.Length > 0)
            {
                result = Request.CreateResponse(HttpStatusCode.OK);
                result.Content = new ByteArrayContent(conteudoArquivo);
                result.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment");
                result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                result.Content.Headers.ContentDisposition.FileName = Uri.EscapeUriString(nomeArquivo);
            }
            else
            {
                result = Request.CreateResponse(HttpStatusCode.Gone);
            }

            return result;
        }

        public async Task<JObject> Post()
        {
            JObject objeto = new JObject();

            string NomeDocumento = "";
            string TipoItem = "";
            string Ambiente = "";
            string Revisao = "";
            string ID_Documento = "";
            
            try
            {
                NomeDocumento = HttpContext.Current.Request.Params["NomeDocumento"];
                TipoItem = HttpContext.Current.Request.Params["TipoItem"];
                Revisao = HttpContext.Current.Request.Params["RevisaoDocumento"];
                Ambiente = HttpContext.Current.Request.Params["Ambiente"];
                ID_Documento = HttpContext.Current.Request.Params["ID_Documento"];

                if (Revisao == null)
                {
                    Revisao = "";
                }
            }
            catch (Exception ex)
            {
                _servidorService.CriarLogErro(ex.Message, "Parâmetros - FilesController - POST", "", 0, ex.StackTrace, "", "", "");
                var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                {
                    Content = new StringContent("Incorrect Parameters")
                };
                throw new HttpResponseException(resp);
            }
            
            long id_tipoitem = 0;
            long id_projeto = 0;
            long idDoc = 0;

            Int64.TryParse(ID_Documento, out idDoc);

            if (!Int64.TryParse(TipoItem, out id_tipoitem) && !Int64.TryParse(Ambiente, out id_projeto))
            {
                var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                {
                    Content = new StringContent("Parameter missing or in incorrect format: TipoItem/Ambiente")
                };
                throw new HttpResponseException(resp);                
            }

            Projeto projeto = _projetoService.RetornarProjeto(id_projeto);
            TipoItem tipoItem = _projetoService.RetornarTipoItem(id_tipoitem);

            Usuario usuario = _usuarioService.RetornarUsuario(Thread.CurrentPrincipal.Identity.Name);

            _triggerService.DefineUsuarioLogado(usuario);
            _documentoService.DefineUsuarioLogado(usuario);
            _bpmService.DefineUsuarioLogado(usuario);
            _arquivoService.DefineUsuarioLogado(usuario);

            Documento documento = new Documento();

            if (idDoc > 0)
            {
                documento = _documentoService.RetornarDocumento(idDoc);
            }
            else if (projeto != null)
            {
                _documentoService.projetoSelecionado = projeto;

                documento = _documentoService.RetornarDocumento(NomeDocumento, projeto.ID_Projeto);
            }
            else if (tipoItem != null)
            {
                projeto = tipoItem.Projeto;
                _documentoService.projetoSelecionado = projeto;

                documento = _documentoService.RetornarDocumento(NomeDocumento, projeto.ID_Projeto, tipoItem.ID_TipoItem, false);               
            }
            else
            {
                var resp = new HttpResponseMessage(HttpStatusCode.NotFound)
                {
                    Content = new StringContent(string.Format("No item type with ID = {0}", id_tipoitem)),
                    ReasonPhrase = "Item Type Not Found"
                };
                throw new HttpResponseException(resp);
            }

            if (documento == null)
            {
                var resp = new HttpResponseMessage(HttpStatusCode.NotFound)
                {
                    Content = new StringContent(string.Format("No document with Name = {0} or ID {1}", NomeDocumento, idDoc)),
                    ReasonPhrase = "Document Not Found"
                };
                throw new HttpResponseException(resp);
            }


            #region UploadArquivos

            string pathTemp = Path.GetTempPath();

            var provider = new MultipartFormDataStreamProvider(pathTemp);

            try
            {
                await Request.Content.ReadAsMultipartAsync(provider);
            }
            catch (Exception ex)
            {
                _servidorService.CriarLogErro(ex.Message, "Erro - FilesController - POST", "", 0, "", "", "", "");

                var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                {
                    Content = new StringContent("Incorrect Parameter: Arquivo")
                };
                throw new HttpResponseException(resp);
            }

            if (provider.FileData.Count > 0)
            {
                try
                {
                    var docfiles = new List<string>();

                    Revisao revSel = null;
                    if (!string.IsNullOrEmpty(Revisao))
                    {
                        revSel = documento.Revisoes/*.Where(r => !r.Excluido)*/.FirstOrDefault(r => r.Numero.ToString() == Revisao);
                    }
                    else
                    {
                        revSel = documento.RevisaoAtual;
                    }
                    var arquivos = revSel.Arquivos.ToList();
                    long ID_ArquivoAtual = 0;

                    //se for multi arquivos limpa a revisao pois não teria como saber qual é o arquivo para atualizar
                    if (documento.TipoItem.PermiteMultiArquivo && arquivos.Count() > 1)
                    {
                        #region Limpa a revisão selecionada e adiciona os novos arquivos                        

                        foreach (Arquivo arquivoAntigo in arquivos)
                        {
                            _arquivoService.RemoverArquivo(documento.ID_Documento, revSel.ID_Revisao, arquivoAntigo.ID_Arquivo);
                        }
                        #endregion
                    }
                    else {
                        if (arquivos.Count() > 0)
                        {
                            ID_ArquivoAtual = arquivos.First().ID_Arquivo;
                        }
                    }

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

                            Revisao rev = null;
                            if (Revisao.Length > 0)
                            {
                                rev = documento.Revisoes/*.Where(r => !r.Excluido)*/.FirstOrDefault(r => r.Numero.ToString() == Revisao);
                            }
                            else
                            {
                                rev = documento.RevisaoAtual;
                            }

                            Arquivo novoArquivo = _arquivoService.CriarArquivoComConteudoRevisao(postedFile.LocalFileName, rev, fileName, usuario.ID_Usuario, ID_ArquivoAtual);
                            
                            try
                            {
                                if (postedFile.Headers.ContentLength > 0)
                                {
                                    novoArquivo.Tamanho = postedFile.Headers.ContentLength;
                                }
                            }
                            catch (ObjectDisposedException ex){}

                            _arquivoService.Save();
                            System.IO.File.Delete(postedFile.LocalFileName);
                        }
                    }
                }
                catch (Exception ex)
                {
                    _servidorService.CriarLogErro(ex.Message, "Erro - FilesController - POST", "", 0, ex.StackTrace, "", "", "");
                    var resp = new HttpResponseMessage(HttpStatusCode.InternalServerError)
                    {
                        Content = new StringContent(ex.Message)
                    };
                    throw new HttpResponseException(resp);
                }
            }
            else
            {
                _servidorService.CriarLogErro("Chamada sem arquivos", "Erro - FilesController - POST", "", 0, "", "", "", "");

                var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                {
                    Content = new StringContent("Incorrect Parameter: Arquivo")
                };
                throw new HttpResponseException(resp);
            }

            #endregion

            objeto["Codigo"] = 0;
            objeto["Mensagem"] = "Arquivo inserido com sucesso.";
            return objeto;
        }

        // PUT: api/Itens/5
        public async Task<JObject> Put()
        {
            JObject objeto = new JObject();

            string NomeDocumento = "";
            string TipoItem = "";
            string Revisao = "";
            string IDArquivo = "";
            string ID_Documento = "";
            bool adicionarNovoArquivo = false;
            string Importador = "";
            string ParteAtualChunk = "";
            string TotalPartesChunk = "";

            try
            {
                NomeDocumento = HttpContext.Current.Request.Params["NomeDocumento"];
                TipoItem = HttpContext.Current.Request.Params["TipoItem"];
                Revisao = HttpContext.Current.Request.Params["RevisaoDocumento"];
                IDArquivo = HttpContext.Current.Request.Params["ID_Arquivo"];
                ID_Documento = HttpContext.Current.Request.Params["ID_Documento"];
                Importador = HttpContext.Current.Request.Params["Importador"];
                ParteAtualChunk = HttpContext.Current.Request.Params["ParteAtualChunk"];
                TotalPartesChunk = HttpContext.Current.Request.Params["TotalPartesChunk"];

                string adicionar = HttpContext.Current.Request.Params["adicionar"];
                bool.TryParse(adicionar, out adicionarNovoArquivo);

            }
            catch (Exception ex)
            {
                _servidorService.CriarLogErro(ex.Message, "Parâmetros - FilesController - PUT", "", 0, ex.StackTrace, "", "", "");
                objeto["Codigo"] = 1;
                objeto["Mensagem"] = "Parâmetros de envio inconsistentes.";
                return objeto;
            }

            try
            {
                Usuario usuario = _usuarioService.RetornarUsuario(Thread.CurrentPrincipal.Identity.Name);
                Documento documento = null;

                long idDoc = 0;
                Int64.TryParse(ID_Documento, out idDoc);
                if (idDoc > 0)
                {
                    documento = _documentoService.RetornarDocumento(idDoc);
                }
                else
                {
                    TipoItem tipoItem = _projetoService.RetornarTipoItem(Convert.ToInt64(TipoItem));

                    Projeto projeto = tipoItem.Projeto;

                    documento = _documentoService.RetornarDocumento(NomeDocumento, projeto.ID_Projeto, tipoItem.ID_TipoItem, false);
                }

                bool importador = false;
                bool.TryParse(Importador, out importador);
                bool importacao = _contaService.UsuarioPossuiPermissao(documento.Projeto, usuario, Usuario.Permissoes.PermiteImportarPlanilha) || importador;

                BPMInstancia instancia = _bpmService.RetornaInstanciaResponsavel(usuario.ID_Usuario, documento.ID_Documento);

                if (instancia == null || (instancia != null && _documentoService.VerificaAcessoInstancia(instancia, usuario)))
                {
                    bool usuarioResponsavel = _documentoService.VerificaResponsabilidade(documento, usuario, instancia.ID_BPMInstancia) && !instancia.Concluido;

                    bool podeAtualizarArquivo = false;
                    if (usuarioResponsavel)
                    {
                        podeAtualizarArquivo = _arquivoService.VerificaPermissaoParaAtualizarArquivo(documento, instancia, null, usuario);
                    }

                    if (!importacao && !podeAtualizarArquivo)
                    {
                        var resp = new HttpResponseMessage(HttpStatusCode.Forbidden);
                        throw new HttpResponseException(resp);
                    }
                }

                #region UploadArquivos

                string pathTemp = Path.GetTempPath();

                var provider = new MultipartFormDataStreamProvider(pathTemp);

                await Request.Content.ReadAsMultipartAsync(provider);

                long ID_ArquivoAtual = 0;
                if (provider.FileData.Count > 0 && documento != null)
                {
                    var docfiles = new List<string>();

                    Revisao revSel = null;
                    if (!string.IsNullOrEmpty(Revisao))
                    {
                        revSel = documento.Revisoes/*.Where(r => !r.Excluido)*/.FirstOrDefault(r => r.Numero.ToString() == Revisao);
                    }
                    else
                    {
                        revSel = documento.RevisaoAtual;
                    }
                    var arquivos = revSel.Arquivos.ToList();
                    

                    if (!string.IsNullOrEmpty(IDArquivo))
                    {
                        ID_ArquivoAtual = Convert.ToInt64(IDArquivo);
                    }

                    //se for multi arquivos limpa a revisao pois não teria como saber qual é o arquivo para atualizar
                    if (documento.TipoItem.PermiteMultiArquivo && arquivos.Count() > 1)
                    {
                        #region Limpa a revisão selecionada e adiciona os novos arquivos
                        
                        if (!adicionarNovoArquivo)
                        {
                            foreach (Arquivo arquivoAntigo in arquivos)
                            {
                                _arquivoService.RemoverArquivo(documento.ID_Documento, revSel.ID_Revisao, arquivoAntigo.ID_Arquivo);
                            }
                        }
                        #endregion
                    }
                    else 
                    {
                        if (arquivos.Count() > 0 && !adicionarNovoArquivo)
                        {
                            ID_ArquivoAtual = arquivos.First().ID_Arquivo;
                        }
                    }
                    
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

                            Revisao rev = null;
                            if (!string.IsNullOrEmpty(Revisao))
                            {
                                rev = documento.Revisoes/*.Where(r => !r.Excluido)*/.FirstOrDefault(r => r.Numero.ToString() == Revisao);
                            }
                            else
                            {
                                rev = documento.RevisaoAtual;
                            }

                            int parte = 0;
                            int total = 1;

                            int.TryParse(ParteAtualChunk, out parte);
                            int.TryParse(TotalPartesChunk, out total);

                            if (importador)
                            {
                                Stream stream = File.OpenRead(postedFile.LocalFileName);
                                Arquivo arq = _arquivoService.CriarArquivoEmChunks(stream, rev, fileName, documento.ID_Projeto, usuario.ID_Usuario, parte, total, ID_ArquivoAtual);
                                ID_ArquivoAtual = arq.ID_Arquivo;
                                stream.Dispose();
                            }
                            else
                            {
                                _arquivoService.CriarArquivoComConteudoRevisao(postedFile.LocalFileName, rev, fileName, usuario.ID_Usuario, ID_ArquivoAtual);
                                System.IO.File.Delete(postedFile.LocalFileName);
                            }
                        }
                    }
                }
                else
                {
                    _servidorService.CriarLogErro("Chamada sem arquivos", "Erro - FilesController - POST", "", 0, "", "", "", "");
                    objeto["Codigo"] = 2;
                    objeto["Mensagem"] = "Arquivo faltante.";
                    return objeto;
                }


                #endregion

                objeto["Codigo"] = 0;
                objeto["Mensagem"] = "Arquivo atualizado com sucesso.";
                objeto["ID_Arquivo"] = ID_ArquivoAtual.ToString();
                return objeto;
            }
            catch (Exception ex)
            {
                _servidorService.CriarLogErro(ex.Message, "Erro - FilesController - PUT", "", 0, ex.StackTrace, "", "", "");
                objeto["Codigo"] = 3;
                objeto["Mensagem"] = "Problemas na atualização do arquivo.";
                return objeto;
            }

        }

        // DELETE: api/Itens/5
        public void Delete(int id)
        {
        }

    }
}
