using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Globalization;
using System.Runtime.InteropServices.ComTypes;
using System.Xml;
using MySqlX.XDevAPI.Common;
using System.Numerics;

namespace MinhaNfe
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual),
 Guid("66EFA8BE-A461-4DD0-84E2-035D87DE07C5")]

    public interface IMinha_Nfe
    {
        dynamic InputMysql(dynamic jsonNotas, string server, string uid, string pwd, string database);

        dynamic QtdXmlEncontrados(string sCNPJ, string dataIni, string dataFim, string login, string senha);

        dynamic ConsultaMinhaNfe(string strperPage, string sCNPJ, string dataIni, string dataFim, string login, string senha);

        string geraData(string data);

        dynamic DownloadXML(string chave, string login, string senha, string dirXml);

        dynamic ManifestoMinhaNfe(string Nfes, string login, string senha, string idEmpresa, string strtpEvento, string xJust, string server, string uid, string pwd, string database);
    }

    [ClassInterface(ClassInterfaceType.None),
    Guid("9B014D4C-D2BA-4405-9709-2985D60EE789")]

    public class Class1 : IMinha_Nfe
    {
        public Class1() { }

        // Formata a data para ser usada na requisição
        public string geraData(string data)
        {
            try
            {
                var dataSplit = data.Split('/');
                var dataFormatada = $"{dataSplit[2]}-{dataSplit[1]}-{dataSplit[0]}";
                return dataFormatada;
            }

            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        // Me retorta a quantidade de nota encontradas dentro do periodo entre a dataIni e dataFim
        public dynamic QtdXmlEncontrados(string sCNPJ, string dataIni, string dataFim, string login, string senha)
        {
            dataIni = geraData(dataIni);
            dataFim = geraData(dataFim);

            var sret = "";
            var host = $"https://api.minhanfe.com.br/nfes?page=1&perPage=10&empresa={sCNPJ}&dataIni={dataIni}T00:00:00.000Z&dataFim={dataFim}T00:00:00.000Z";
            

            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
                    | SecurityProtocolType.Tls11
                    | SecurityProtocolType.Tls12
                    | SecurityProtocolType.Ssl3;

                var request1 = (HttpWebRequest)HttpWebRequest.Create($"{host}");

                var base64authorization = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{login}:{senha}"));

                request1.Method = "GET";
                request1.Headers.Add("Authorization", "Basic " + base64authorization);

                using (var response = (HttpWebResponse)request1.GetResponse())
                {

                    int status = (int)response.StatusCode;
                    if (status == 200)
                    {
                        var streamReader = new StreamReader(response.GetResponseStream());
                        var result = streamReader.ReadToEnd();
                        var sval1 = result.ToString();
                        dynamic jsonret = Newtonsoft.Json.JsonConvert.DeserializeObject(result);
                        int qtdXmlEncontrados = jsonret.count;

                        return qtdXmlEncontrados.ToString();
                    }//if status eq 200
                    else
                    {
                        sret = response.StatusDescription.ToString();
                        return sret;
                    }//else status 200
                }//var response
            }//try
            catch (Exception ex)
            {
                return ex.ToString();
            }//catch
        }

        // Consulta A api do MinhaNfe com a quantidade de notas encontradas no periodo selecionado 
        public  dynamic ConsultaMinhaNfe(string strperPage, string sCNPJ, string dataIni, string dataFim, string login, string senha)
        {

            dataIni = geraData(dataIni);
            dataFim = geraData(dataFim);

            var sret = "";
            var host = $"https://api.minhanfe.com.br/nfes?page=1&perPage={strperPage}&empresa={sCNPJ}&dataIni={dataIni}T00:00:00.000Z&dataFim={dataFim}T23:59:00.000Z";
            
            try
            {
                var request1 = (HttpWebRequest)HttpWebRequest.Create($"{host}");

                var base64authorization = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{login}:{senha}"));

                request1.Method = "GET";
                request1.Headers.Add("Authorization", "Basic " + base64authorization);

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
                    | SecurityProtocolType.Tls11
                    | SecurityProtocolType.Tls12
                    | SecurityProtocolType.Ssl3;

                using (var response = (HttpWebResponse)request1.GetResponse())
                {

                    int status = (int)response.StatusCode;
                    if (status == 200)
                    {
                        var streamReader = new StreamReader(response.GetResponseStream());
                        var result = streamReader.ReadToEnd();
                        var sval1 = result.ToString();
                        dynamic jsonret = Newtonsoft.Json.JsonConvert.DeserializeObject(result);



                        return jsonret;//retona objeto completo com todas as notas encotradas


                    }//if status 200
                    else
                    {
                        sret = response.StatusDescription.ToString();
                        return sret;
                    }//else status 200

                }//using var response
            }//try requisição
            catch (Exception ex)
            {
                return ex.ToString() ;
            }//catch requisição
        }//consultaNfe


        // Alimenta o banco de dados local com as notas encontradas na consulta ao banco de dados do MinhaNfe
        public dynamic InputMysql(dynamic jsonNotas, string server, string uid, string pwd, string database)
        {
            try
            {
                
                var notas = jsonNotas.notas;
                int iCont = notas.Count;

                int iDuplicados = 0;

                int i = 0;
                while (i < iCont)
                {
                    var date = notas[i].dhEmi;
                    var formattedDate = String.Format("{0:yyyy-MM-dd}", date);
                    var Cancelada = "";
                    var nvalor = notas[i].valor.ToString().Replace(",", ".");

                    string chave = Convert.ToString(notas[i].chave);
                    string numero = chave.Substring(25, 9);

                    bool bCancel = notas[i].cancelada;
                    if (bCancel)
                    {
                        Cancelada = "S";
                    }//if bCancel
                    else
                    {
                        Cancelada = "N";
                    }//else bCancel eq true

                    try
                    {
                        var strConexao = $"server={server};uid={uid};pwd={pwd};database={database}";
                        var conexao = new MySqlConnection(strConexao);
                        conexao.Open();

                        var strCommando = new MySqlCommand($"INSERT INTO sqxmltab(CHAVE, EMITENTE, DESTINATARIO, VALOR, DATAEMI, CANCELADA, CNPJ_DEST, NUMERO, MANIFESTO) VALUES('{chave}','{Convert.ToString(notas[i].emitente.nome)}'," +
                           $"'{Convert.ToString(notas[i].empresa.nome)}',{nvalor},'{formattedDate}','{Cancelada}'," +
                           $"'{Convert.ToString(notas[i].empresa.cnpj)}',{numero},'{notas[i].manifesto}');", conexao);
                        var reader = strCommando.ExecuteReader();
                    }//try for
                    catch
                    {
                        iDuplicados++;
                    }//catch for

                    i++;
                }//while

                if (iCont - iDuplicados == 0)
                {
                    return new string[] { "204", "Nenhum item novo a ser inserido" };//cod da requisão e msg erro
                }
                return new string[] { "200", $"{iCont - iDuplicados} => Dados inseridos corretamente.\n{iDuplicados} => Dados duplicados não inseridos" };



            }//try
            catch (Exception ex)
            {
                return new string[] { "500", ex.ToString() };
            }//catch

        }//input mysql


        // Faz o download do XML especificado
        public dynamic DownloadXML(string chave, string login, string senha, string dirXml)
        {
            var sret = "";
            var host = $"https://api.minhanfe.com.br/nfes/{chave}";

            try
            {
                var request1 = (HttpWebRequest)HttpWebRequest.Create($"{host}");

                var base64authorization = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{login}:{senha}"));

                request1.Method = "GET";
                request1.Headers.Add("Authorization", "Basic " + base64authorization);

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
                    | SecurityProtocolType.Tls11
                    | SecurityProtocolType.Tls12
                    | SecurityProtocolType.Ssl3;

                using (var response = (HttpWebResponse)request1.GetResponse())
                {

                    int status = (int)response.StatusCode;
                    if (status == 200)
                    {
                        var streamReader = new StreamReader(response.GetResponseStream());
                        var result = streamReader.ReadToEnd();
                        var sval1 = result.ToString();
                        XmlDocument xmlDocument = new XmlDocument();
                        xmlDocument.LoadXml(sval1);
                        xmlDocument.Save($"{dirXml}{chave}.xml");//salvar xml no destino para leitura


                        return status.ToString();//200


                    }//if status 200
                    else
                    {
                        sret = response.StatusDescription.ToString();
                        return sret;
                    }//else status 200

                }//using var response
            }//try requisição
            catch (Exception ex)
            {
                return ex.ToString();
            }//catch requisição
        }//DownlaodXml


        // Manifesto de destinatario para diversar notas fiscais simultaneas: NÃO UTILIZADO AINDA
        public dynamic multiManifestoMinhaNfe(string Nfes, string login, string senha, string idEmpresa, string strtpEvento, string xJust, string server, string uid, string pwd, string database)
        {
            var sret = "";
            var host = $"https://api.minhanfe.com.br/api/manifesto";
            int tpEvento = Convert.ToInt32(strtpEvento);
            var descricaoManifesto = "";

            switch (tpEvento)
            {
                case 210200:
                    descricaoManifesto = "Confirmação da Operação";
                    break;
                case 210220:
                    descricaoManifesto = "Desconhecimento da Operação";
                    break;
                case 210240:
                    descricaoManifesto = "Operação não Realizada";
                    break;
            }

            try
            {
                var request1 = (HttpWebRequest)HttpWebRequest.Create($"{host}");

                var base64authorization = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{login}:{senha}"));

                request1.Method = "POST";
                request1.Headers.Add("Authorization", "Basic " + base64authorization);
                request1.ContentType = "application/json";

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
                    | SecurityProtocolType.Tls11
                    | SecurityProtocolType.Tls12
                    | SecurityProtocolType.Ssl3;

                try
                {

                    using (var streamWriter = new StreamWriter(request1.GetRequestStream()))
                    {
                        
                        var json = (new
                        {
                            idEmpresa = $"{idEmpresa}",
                            NFes = new List<string>() {$"{Nfes}"},
                            tpEvento = $"{tpEvento}",
                            xJust = $"{xJust}"
                        });

                        var jsontk = JsonConvert.SerializeObject(json);

                        streamWriter.Write(jsontk);
                    }



                }//try json
                catch (Exception ex)
                {

                    return ex.Message;

                }//catch json

                using (var response = (HttpWebResponse)request1.GetResponse())
                {

                    int status = (int)response.StatusCode;
                    if (status == 200)
                    {
                        var streamReader = new StreamReader(response.GetResponseStream());
                        var result = streamReader.ReadToEnd();
                        var sval1 = result.ToString();
                        dynamic jsonret = Newtonsoft.Json.JsonConvert.DeserializeObject(result);
                        var qtdNfes = jsonret.eventos.Count;

                        try
                        {
                            for (int i = 0; i < qtdNfes; i++)
                            {
                                var chave = Convert.ToString(jsonret.eventos[i].chNFe);
                                var xMotivo = Convert.ToString(jsonret.eventos[i].chNFe);

                                var strConexao = $"server={server};uid={uid};pwd={pwd};database={database}";
                                var conexao = new MySqlConnection(strConexao);
                                conexao.Open();

                                var strCommando = new MySqlCommand($"UPDATE sqxmltab SET MANIFESTO = '{descricaoManifesto}' WHERE CHAVE = '{chave}';", conexao);
                                var reader = strCommando.ExecuteReader();

                            }

                            return "";
                        }catch(Exception ex) { return ex.ToString(); }
                        


                        return status.ToString();//200


                    }//if status 200
                    else
                    {
                        sret = response.StatusDescription.ToString();
                        return sret;
                    }//else status 200

                }//using var response
            }//try requisição
            catch (Exception ex)
            {
                return ex.ToString();
            }//catch requisição
        }//MULTImanifestoMinhaNfe

        // Manifesto de destinatario pela api do MinhaNfe para uma unica nota fiscal selecionada
        public dynamic ManifestoMinhaNfe(string Nfes, string login, string senha, string idEmpresa, string strtpEvento, string xJust, string server, string uid, string pwd, string database)
        {
            var sret = "";
            var host = $"https://api.minhanfe.com.br/api/manifesto";
            int tpEvento = Convert.ToInt32(strtpEvento);
            var descricaoManifesto = "";

            switch (tpEvento)
            {
                case 210200:
                    descricaoManifesto = "Confirmacao da Operacao";
                    break;
                case 210220:
                    descricaoManifesto = "Desconhecimento da Operacao";
                    break;
                case 210240:
                    descricaoManifesto = "Operacao nao Realizada";
                    break;
            }

            try
            {
                var request1 = (HttpWebRequest)HttpWebRequest.Create($"{host}");

                var base64authorization = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{login}:{senha}"));

                request1.Method = "POST";
                request1.Headers.Add("Authorization", "Basic " + base64authorization);
                request1.ContentType = "application/json";

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
                    | SecurityProtocolType.Tls11
                    | SecurityProtocolType.Tls12
                    | SecurityProtocolType.Ssl3;

                try
                {

                    using (var streamWriter = new StreamWriter(request1.GetRequestStream()))
                    {

                        var json = (new
                        {
                            idEmpresa = $"{idEmpresa}",
                            NFes = new List<string>() { $"{Nfes}" },
                            tpEvento = tpEvento,
                            xJust = $"{xJust}"
                        });

                        var jsontk = JsonConvert.SerializeObject(json);

                        streamWriter.Write(jsontk);
                    }



                }//try json
                catch (Exception ex)
                {

                    return ex.Message;

                }//catch json

                using (var response = (HttpWebResponse)request1.GetResponse())
                {

                    int status = (int)response.StatusCode;
                    if (status == 200)
                    {
                        var streamReader = new StreamReader(response.GetResponseStream());
                        var result = streamReader.ReadToEnd();
                        var sval1 = result.ToString();
                        dynamic jsonret = Newtonsoft.Json.JsonConvert.DeserializeObject(result);

                        var cStat = Convert.ToString(jsonret.eventos[0].cStat);
                        var xMotivo = Convert.ToString(jsonret.eventos[0].xMotivo);

                        if (cStat == "135")
                        {
                            try
                            {
                                var chave = Convert.ToString(jsonret.eventos[0].chNFe);

                                var strConexao = $"server={server};uid={uid};pwd={pwd};database={database}";
                                var conexao = new MySqlConnection(strConexao);
                                conexao.Open();

                                var strCommando = new MySqlCommand($"UPDATE sqxmltab SET MANIFESTO = '{descricaoManifesto}' WHERE CHAVE = '{chave}';", conexao);
                                var reader = strCommando.ExecuteReader();

                                return xMotivo;

                            }
                            catch (Exception ex) { return ex.ToString(); }

                        }//if cStat == 135
                        else
                        {
                            return  xMotivo;

                        }//else cStat == 135

                    }//if status 200
                    else
                    {
                        sret = response.StatusDescription.ToString();
                        return sret;
                    }//else status 200

                }//using var response
            }//try requisição
            catch (Exception ex)
            {
                return ex.ToString();
            }//catch requisição
        }//manifestoMinhaNfe




    }//Class1

    
}
