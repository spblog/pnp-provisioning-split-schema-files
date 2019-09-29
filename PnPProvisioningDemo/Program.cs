using System;
using System.Configuration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;

namespace PnPProvisioningDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var ctx = GetContext())
            {
                SimpleProvision(ctx);
            }
        }

        public static ClientContext GetContext()
        {
            var clientId = ConfigurationManager.AppSettings.Get("ClientId");
            var clientSecret = ConfigurationManager.AppSettings.Get("ClientSecret");
            var spUrl = ConfigurationManager.AppSettings.Get("SPUrl");

            var mngr = new AuthenticationManager();
            var appOnlyCtx = mngr.GetAppOnlyAuthenticatedContext(spUrl, clientId, clientSecret);

            return appOnlyCtx;
        }

        public static void SimpleProvision(ClientContext ctx)
        {

            var provider = new XMLFileSystemTemplateProvider($@"{Environment.CurrentDirectory}\..\..\templates\awesome-team\", "");

            var template = provider.GetTemplate("awesome-team.xml");

            var connector = new FileSystemConnector($@"{Environment.CurrentDirectory}\..\..\templates\awesome-team\files", "");

            template.Connector = connector;

            var ptai = new ProvisioningTemplateApplyingInformation
            {
                ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                {
                    Console.WriteLine("{0:00}/{1:00} - {2}", progress, total, message);
                }
            };

            ctx.Web.ApplyProvisioningTemplate(template, ptai);

            provider.SaveAs(template, "awesome-team.generated.xml");
        }
    }
}
