using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Security.Authentication.ExtendedProtection;
using System.Security.Principal;
using System.Text;
using System.Web;
using System.Web.Configuration;
using System.Web.Mvc;
using System.Web.Routing;
using UCT.Controllers;
using UCT.Models;
using UCT.UnitTests.Model;

namespace UCT.UnitTests.Controllers
{
    [TestClass]
    public class CompetencyControllerTest
    {
        private const int PROGRAM_ID = 59;

        private static CompetencyController GetCompetencyController(IUCTRepository repository, NameValueCollection routingValues)
        {
            var mockContext = new MockHttpContext();
            CompetencyController controller = new CompetencyController(repository, mockContext.User);

            RouteData routeData = new RouteData();
            if ((routingValues != null) && (routingValues.Count > 0))
            {
                foreach (string key in routingValues)
                    routeData.Values.Add(key, routingValues[key]);
            }

            controller.ControllerContext = new ControllerContext()
            {
                Controller = controller,
                RequestContext = new RequestContext(mockContext, routeData)
            };

            return controller;
        }

        public interface IHttpObject { }

        public class FakeHttpRequest : HttpRequestBase, IHttpObject
        {
            private readonly NameValueCollection form;
            private readonly NameValueCollection queryString;
            private readonly NameValueCollection headers;
            private readonly NameValueCollection serverVariables;
            private readonly HttpCookieCollection cookies;
            private HttpFileCollectionBase files;
            private Uri url;
            private string method;
            private bool isLocal;
            private string userHostAddress;
            private string applicationPath;
            private string[] acceptTypes;
            private Stream inputStream;
            private bool isAuthenticated;

            private string anonymousId;
            private string appRelativeCurrentExecutionFilePath;
            private HttpBrowserCapabilitiesBase browser = new HttpBrowserCapabilitiesWrapper(BrowserCapabilities.GetHttpBrowserCapabilities(null, "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.2; WOW64; Trident/6.0)"));
            private ChannelBinding httpChannelBinding;
            private HttpClientCertificate clientCertificate;
            private int contentLength;
            private string currentExecutionFilePath;
            private Stream filter;
            private bool isSecureConnection;
            private WindowsIdentity logonUserIdentity;
            private NameValueCollection @params;
            private string physicalApplicationPath;
            private string physicalPath;
            private RequestContext requestContext;
            private int totalBytes;
            private Uri urlReferrer;
            private string userAgent = "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.2; WOW64; Trident/6.0)";
            private string[] userLanguages;
            private string userHostName;
            private string pathInfo;

            private static NameValueCollection ParseQueryString(string url)
            {
                return HttpUtility.ParseQueryString(url);
            }

            public FakeHttpRequest(Uri url = null, string method = "GET")
            {
                this.url = url ?? new Uri("http://localhost");
                this.method = method;
                acceptTypes = new string[] { };
                queryString = ParseQueryString(this.url.Query);
                form = new NameValueCollection();
                headers = new NameValueCollection();
                serverVariables = new NameValueCollection();
                cookies = new HttpCookieCollection();
            }

            public FakeHttpRequest SetUrl(Uri url)
            {
                this.url = url;
                queryString.Clear();
                queryString.Add(ParseQueryString(url.Query));
                return this;
            }

            public override string this[string key]
            {
                get { return new NameValueCollection { form, queryString }[key]; }
            }

            public override bool IsAuthenticated { get { return isAuthenticated; } }
            public override Uri Url { get { return url; } }
            public override bool IsLocal { get { return isLocal; } }
            public override string ApplicationPath { get { return applicationPath; } }
            public override string HttpMethod { get { return method; } }
            public override string UserHostAddress { get { return userHostAddress; } }
            public override string[] AcceptTypes { get { return acceptTypes; } }
            public override string RequestType { get; set; }
            public override string ContentType { get; set; }
            public override Encoding ContentEncoding { get; set; }
            public override void ValidateInput() { }
            public override string RawUrl { get { return url.PathAndQuery; } }
            public override NameValueCollection Form { get { return form; } }
            public override NameValueCollection QueryString { get { return queryString; } }
            public override NameValueCollection Headers { get { return headers; } }
            public override NameValueCollection ServerVariables { get { return serverVariables; } }
            public override HttpCookieCollection Cookies { get { return cookies; } }
            public override HttpFileCollectionBase Files { get { return files; } }
            public override string Path { get { return Url.AbsolutePath; } }
            public override string FilePath { get { return Url.AbsolutePath; } }
            public override string PathInfo { get { return pathInfo; } }
            public override Stream InputStream { get { return inputStream; } }
            public override string AnonymousID { get { return anonymousId; } }
            public override string AppRelativeCurrentExecutionFilePath { get { return appRelativeCurrentExecutionFilePath; } }
            public override HttpBrowserCapabilitiesBase Browser { get { return browser; } }
            public override ChannelBinding HttpChannelBinding { get { return httpChannelBinding; } }
            public override HttpClientCertificate ClientCertificate { get { return clientCertificate; } }
            public override int ContentLength { get { return contentLength; } }
            public override string CurrentExecutionFilePath { get { return currentExecutionFilePath; } }
            public override Stream Filter { get { return filter; } set { filter = value; } }
            public override bool IsSecureConnection { get { return isSecureConnection; } }
            public override WindowsIdentity LogonUserIdentity { get { return logonUserIdentity; } }
            public override NameValueCollection Params { get { return @params; } }
            public override string PhysicalApplicationPath { get { return physicalApplicationPath; } }
            public override string PhysicalPath { get { return physicalPath; } }
            public override RequestContext RequestContext { get { return requestContext; } }
            public override int TotalBytes { get { return totalBytes; } }
            public override Uri UrlReferrer { get { return urlReferrer; } }
            public override string UserAgent { get { return userAgent; } }
            public override string[] UserLanguages { get { return userLanguages; } }
            public override string UserHostName { get { return userHostName; } }
        }


        private class MockHttpContext : HttpContextBase
        {
            private readonly IPrincipal _user = new GenericPrincipal(new GenericIdentity("program"), new string[] { "ProgramDirector" } );//null /* roles */
            private IDictionary _items = new Dictionary<object, object>();
            private FakeHttpResponse _response = new FakeHttpResponse();
            private FakeHttpRequest _request = new FakeHttpRequest();
            
            public override IPrincipal User
            {
                get
                {
                    return _user;
                }
                set
                {
                    base.User = value;
                }
            }

            public override IDictionary Items
            {
                get
                {
                    return _items;
                }
            }
            public override HttpResponseBase Response
            {
                get
                {
                    return _response;
                }
            }
            public override HttpRequestBase Request
            {
                get
                {
                    return _request;
                }
            }
        }

        public class FakeHttpResponse : HttpResponseBase
        {
            private HttpCookieCollection _cookies = new HttpCookieCollection();
            
            public override void Redirect(string url)
            {
                RedirectLocation = url;
            }

            public override string RedirectLocation
            {
                get;
                set;
            }
            public override HttpCookieCollection Cookies
            {
                get
                {
                    return _cookies;
                }
            }
        }

        public class BrowserCapabilities
        {
            public static HttpBrowserCapabilities GetHttpBrowserCapabilities(NameValueCollection headers, string userAgent)
            {
                var factory = new BrowserCapabilitiesFactory();
                var browserCaps = new HttpBrowserCapabilities();
                var hashtable = new Hashtable(180, StringComparer.OrdinalIgnoreCase);
                hashtable[string.Empty] = userAgent;
                browserCaps.Capabilities = hashtable;
                factory.ConfigureBrowserCapabilities(headers, browserCaps);
                factory.ConfigureCustomCapabilities(headers, browserCaps);
                return browserCaps;
            }
        }

        [TestMethod]
        public void Index_Get_AsksForIndexView()
        {            
            // Arrange
            CompetencyController controller = GetCompetencyController(new InMemoryUCTRepository(), null);
            
            // Act
            ViewResult result = (ViewResult)controller.Index(null);
            // Assert
            Assert.AreEqual("Index", result.ViewName);            
        }

        [TestMethod]
        public void LoadCreateLearningGoalForm_Post_GetJson()
        {
            NameValueCollection routingParams =  new NameValueCollection();
            routingParams.Add("controller", "Competency");
            routingParams.Add("action", "LoadCreateLearningGoal");

            //using(HttpSimulator simulator = new HttpSimulator())
            //{
                // Arrange
                CompetencyController controller = GetCompetencyController(new InMemoryUCTRepository(), routingParams);
            
                // Simply executing a method during a unit test does just that - executes a method, and no more. 
                // The MVC pipeline doesn't run, so binding and validation don't run.
                //controller.ModelState.AddModelError("", "mock error message");
                //Contact model = GetContactNamed(1, "", "");

                // Act
                var result = (JsonResult)controller.LoadCreateLearningGoal(PROGRAM_ID);

                // Assert
                Assert.IsNotNull(result.Data);
            //}

            // Assert
            //var model = (IEnumerable<Contact>)result.ViewData.Model;
            //CollectionAssert.Contains(model.ToList(), contact1);
            //CollectionAssert.Contains(model.ToList(), contact1);
        }

        [TestMethod]
        public void ExportReport()
        {
            NameValueCollection routingParams = new NameValueCollection();
            routingParams.Add("controller", "Competency");
            routingParams.Add("action", "LoadCreateLearningGoal");

            //using(HttpSimulator simulator = new HttpSimulator())
            //{
            // Arrange
            CompetencyController controller = GetCompetencyController(new InMemoryUCTRepository(), routingParams);

            // Simply executing a method during a unit test does just that - executes a method, and no more. 
            // The MVC pipeline doesn't run, so binding and validation don't run.
            //controller.ModelState.AddModelError("", "mock error message");
            //Contact model = GetContactNamed(1, "", "");

            // Act
            var result = (JsonResult)controller.Export(PROGRAM_ID);

            // Assert
            Assert.IsNotNull(result.Data);
            //}

            // Assert
            //var model = (IEnumerable<Contact>)result.ViewData.Model;
            //CollectionAssert.Contains(model.ToList(), contact1);
            //CollectionAssert.Contains(model.ToList(), contact1);
        }
    }
}
