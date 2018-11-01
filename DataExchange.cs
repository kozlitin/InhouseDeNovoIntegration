using System;
using System.Net;
using System.Runtime.InteropServices;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;

namespace InhouseDeNovoIntegration
{
    [Guid("1F8CB412-E62C-40D7-B494-A469AB4102C4")]
    public interface IDeNovoIntegration
    {
        [DispId(1)]
        string TryTerrasoftLogin(string userName, string userPassword, string _baseUri);
        string GetTerrasoftBillingActs(string Month, string Year);
    }

    [Guid("172346A4-D66A-4823-845E-3133998631BC"),
     ClassInterface(ClassInterfaceType.None),
     ComSourceInterfaces(typeof(IDeNovoIntegration))]
    public class DataExchange : IDeNovoIntegration
    {

        // HTTP-адрес приложения Terrasoft.
        private static string baseUri = "https://devbpm.de-novo.biz";
        // Контейнер для Cookie аутентификации bpm'online. Необходимо использовать в последующих запросах.
        // Это самый важный результирующий объект, для формирования свойств которого разработана
        // вся остальная функциональность примера.
        public static CookieContainer AuthCookie = new CookieContainer();
        // Строка запроса к методу Login сервиса AuthService.svc.
        private static string authServiceUri = @"/ServiceModel/AuthService.svc/Login";
        private static string getBillingActsServiceUri = @"/0/rest/InhouseDeNovoDataExport/GetBillingActs?";

        // Выполняет запрос на аутентификацию пользователя.
        public string TryTerrasoftLogin(string userName, string userPassword, string _baseUri)
        {
            baseUri = _baseUri;
            // Создание экземпляра запроса к сервису аутентификации.
            var authRequest = HttpWebRequest.Create(baseUri + authServiceUri) as HttpWebRequest;
            // Определение метода запроса.
            authRequest.Method = "POST";
            // Определение типа контента запроса.
            authRequest.ContentType = "application/json";
            // Включение использования cookie в запросе.
            authRequest.CookieContainer = AuthCookie;

            authRequest.ServerCertificateValidationCallback = delegate (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };

            // Помещение в тело запроса учетной информации пользователя.
            using (var requestStream = authRequest.GetRequestStream())
            {
                using (var writer = new StreamWriter(requestStream))
                {
                    writer.Write(@"{
                    ""UserName"":""" + userName + @""",
                    ""UserPassword"":""" + userPassword + @"""
                    }");
                }
            }

            string responseText = "";
            // Получение ответа от сервера. Если аутентификация проходит успешно, в свойство AuthCookie будут
            // помещены cookie, которые могут быть использованы для последующих запросов.
            using (var response = (HttpWebResponse)authRequest.GetResponse())
            {
                using (var reader = new StreamReader(response.GetResponseStream()))
                {
                    responseText = reader.ReadToEnd();
                }

            }

            return responseText;
        }

        public string GetTerrasoftBillingActs(string Month, string Year)
        {
            var getContactIdRequest = HttpWebRequest.Create(baseUri + getBillingActsServiceUri + "Month=" + Month + "&Year=" + Year) as HttpWebRequest;
            getContactIdRequest.Method = "GET";
            getContactIdRequest.ContentType = "application/json";
            getContactIdRequest.CookieContainer = AuthCookie;

            getContactIdRequest.ServerCertificateValidationCallback = delegate (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };

            string responseText;

            using (var response = (HttpWebResponse)getContactIdRequest.GetResponse())
            {
                using (var reader = new StreamReader(response.GetResponseStream()))
                {
                    responseText = reader.ReadToEnd();
                }

            }

            return responseText;
        }

    }

}