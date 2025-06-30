using GreenDoc.Data;
using GreenDoc.Infrastructure;
using GreenDoc.Models;
using GreenDoc.Models.API;
using GreenDoc.Services;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Web;
using System.Web.Http;
using System.Web.Security;

namespace GreenDoc.Web.Api.Controllers
{
    [MyBasicAuthenticationFilter]
    public class ComentsController : ApiController
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
        public ComentarioService _comentarioService;
        public TriggerService _triggerService;

        public ComentsController()
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
            _comentarioService = new ComentarioService(unitOfWork, _triggerService);

        }

        // GET: api/Contas
        public JArray Get()
        {
            string Revisao = HttpContext.Current.Request.Params["Revisao"];

            long id_Revisao = 0;
            if (!Int64.TryParse(Revisao, out id_Revisao))
            {
                var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                {
                    Content = new StringContent("Parameter was in incorrect format: Revisao")
                };
                throw new HttpResponseException(resp);
            }

            List<Comentario> coments = null;

            Revisao revisao = _documentoService.RetornaRevisao(id_Revisao);

            if (revisao != null)
            {
                if (!VerifacaPermissaoDocumento(Thread.CurrentPrincipal.Identity.Name, revisao.Documento.ID_Documento))
                {
                    throw new HttpResponseException(System.Net.HttpStatusCode.Unauthorized);
                }
                else
                {
                    coments = _comentarioService.ListarComentarios(revisao).ToList();
                }
            }
            else
            {
                throw new HttpResponseException(System.Net.HttpStatusCode.NoContent);
            }

            JArray array = new JArray();

            foreach (var coment in coments)
            {
                JObject objeto = new JObject();

                objeto["ID_Comentario"] = coment.ID_Comentario;
                objeto["ID_Revisao"] = coment.ID_Revisao;
                objeto["Descricao"] = coment.Descricao;
                objeto["Data"] = coment.Data;
                objeto["Usuario"] = coment.Usuario.Nome;

                objeto["NomeAnexo"] = "";
                objeto["Anexo"] = "";              

                if (coment.AnexosComentarios.Count == 1)
                {
                    var anexoComentario = coment.AnexosComentarios.First();

                    byte[] conteudoArquivo = _arquivoService.RetornarConteudoArquivo(anexoComentario.Arquivo);
                    if (conteudoArquivo != null && conteudoArquivo.Length > 0)
                    {
                        objeto["NomeAnexo"] = Uri.EscapeUriString(anexoComentario.Arquivo.Nome + anexoComentario.Arquivo.Extensao);
                        objeto["Anexo"] = Convert.ToBase64String(conteudoArquivo);
                    }
                }
                else if (coment.AnexosComentarios.Count > 1)
                {
                    string nomeArquivoZip = "Anexos_" + string.Join("_", coment.AnexosComentarios.Select(ac => ac.ID_Anexo).ToArray()) + ".zip";

                    using (var compressedFile = new MemoryStream())
                    {
                        using (var zipArchive = new ZipArchive(compressedFile, ZipArchiveMode.Update, false))
                        {
                            foreach (AnexoComentario anexoComentario in coment.AnexosComentarios)
                            {
                                Arquivo arquivo = anexoComentario.Arquivo;

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
                        objeto["NomeAnexo"] = nomeArquivoZip;
                        objeto["Anexo"] = Convert.ToBase64String(compressedFile.ToArray());
                    }
                }

                array.Add(objeto);
            }

            return array;
        }

        // GET: api/Itens/5
        public string Get(int id)
        {
            return "value";
        }

        public HttpResponseMessage Post()
        {   
            bool publico;
            var httpRequest = HttpContext.Current.Request;
            string paramPublico = HttpContext.Current.Request.Params["Publico"];
            string NomeDocumento = HttpContext.Current.Request.Params["NomeDocumento"];
            string TipoItem = HttpContext.Current.Request.Params["TipoItem"];
            string comentario = HttpContext.Current.Request.Params["Comentario"];
            string usuarioComentario = HttpContext.Current.Request.Params["UsuarioComentario"];
            var resp = new HttpResponseMessage(HttpStatusCode.BadRequest) { Content = new StringContent("Parameter in incorrect format: Publico") };

            TipoItem tipoItem = _projetoService.RetornarTipoItem(Convert.ToInt64(TipoItem));

            Projeto projeto = tipoItem.Projeto;
            _documentoService.projetoSelecionado = projeto;

            Usuario _usuario = _usuarioService.RetornarUsuario(usuarioComentario);
            _documentoService.DefineUsuarioLogado(_usuario);
            _projetoService.projetoSelecionado = projeto;

            Documento documento = _documentoService.RetornarDocumento(NomeDocumento, projeto.ID_Projeto, tipoItem.ID_TipoItem, false);           

            if (!VerifacaPermissaoDocumento(usuarioComentario, documento.ID_Documento))
            {
                throw new HttpResponseException(System.Net.HttpStatusCode.Unauthorized);
            }

            if (string.IsNullOrEmpty(paramPublico))
            {
                publico = true;
            }
            else
            {
                if(!bool.TryParse(paramPublico, out publico))
                {
                    throw new HttpResponseException(resp);
                }
            }

            Comentario coment = _comentarioService.CriarComentario(comentario, _usuario.ID_Usuario, (long)documento.ID_RevisaoAtual, publico, null);

            #region UploadArquivos

            string pathTemp = Path.GetTempPath();

            if (httpRequest.Files.Count > 0)
            {
                var docfiles = new List<string>();

                foreach (string file in httpRequest.Files)
                {

                    var postedFile = httpRequest.Files[file];

                    if (postedFile != null)
                    {
                        if (postedFile.ContentLength > 0)
                        {
                            Arquivo arquivo = _arquivoService.CriarArquivo(postedFile.FileName, projeto.ID_Projeto, _usuario.ID_Usuario);

                            _arquivoService.AppendConteudoArquivo(arquivo, postedFile.InputStream);

                            _arquivoService.AdicionaArquivoComentario(coment, arquivo);
                        }
                    }
                }
            }
            
            return Request.CreateResponse(HttpStatusCode.OK);
            #endregion

        }

        // PUT: api/Itens/5
        public void Put()
        {


        }

        // DELETE: api/Itens/5
        public void Delete(int id)
        {
        }

        private bool VerifacaPermissaoDocumento(string usuario, long idDocumento)
        {
            Usuario _usuario = _usuarioService.RetornarUsuario(usuario);

            BPMInstancia instancia = _bpmService.RetornaInstanciaResponsavel(_usuario.ID_Usuario, idDocumento);

            if (_usuario != null && _usuario.Services == true)
                return true;

            if (usuario == null || (instancia != null && !_documentoService.VerificaAcessoInstancia(instancia, _usuario)))
            {
                return false;
            }
            else
            {
                return true;
            }
        }
    }
}