using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Publishing;
using OfficeDevPnP.Core.Utilities.Themes;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Graph.Model;
using System.Runtime.Remoting.Contexts;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.V201903;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.V201807;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.V201909;
using Microsoft.SharePoint.Navigation;
using System.Net;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration.Navigation;
using System.Xml.Linq;
using OfficeDevPnP.Core.Extensions;
using System.Web.Mvc;
using OfficeDevPnP.Core.Pages;
using System.IO;
using static System.Collections.Specialized.BitVector32;
using CanvasSection = OfficeDevPnP.Core.Pages.CanvasSection;
using CanvasControl = OfficeDevPnP.Core.Pages.CanvasControl;
using Newtonsoft.Json.Linq;

namespace TestesCSOM
{
  internal class Program
  {
    static void Main(string[] args)
    {

      //string password = "Jo@o2307";
      //string username = "Joao.Neto@class-solutions.com.br";
      //ClientContext context = new ClientContext("https://classsolutions.sharepoint.com/sites/TestesLang");
      //SecureString SecurePassword = new SecureString();
      //foreach (char c in password.ToCharArray())
      //{
      //  SecurePassword.AppendChar(c);
      //}
      ////NetworkCredential _myCredentials = new NetworkCredential(username, password);
      //context.Credentials = new SharePointOnlineCredentials(username, SecurePassword);
      ////Web web = context.Web;
      ////context.Load(web);
      ////context.ExecuteQuery();
      ////ClientSidePageLayoutType layout = ClientSidePageLayoutType.Spaces;
      //var web = context.Web;
      //context.Load(web);
      //context.ExecuteQuery();

      //OfficeDevPnP.Core.Pages.ClientSidePage page = OfficeDevPnP.Core.Pages.ClientSidePage.Load(context, "Home.aspx");
      //page.LayoutType = ClientSidePageLayoutType.Home;
      ////IEnumerable<ClientSideComponent> components = page.AvailableClientSideComponents();

      ////foreach (CanvasSection section in page.Sections)
      ////{
      ////  foreach (CanvasControl control in section.Controls)
      ////  {
      ////    control.Delete();
      ////  }
      ////}

      ////page.Sections.Clear();

      ////page.Save();
      ////page.Publish();


      //// Verifica se a página tem seções
      //if (page.Sections.Count > 0)
      //{
      //  // Loop pelas seções da página
      //  foreach (CanvasSection section in page.Sections.ToList())
      //  {
      //    // Loop pelas webparts em cada seção
      //    foreach (CanvasControl control in section.Controls.ToList())
      //    {
      //      // Exclui a webpart
      //      control.Delete();
      //    }
      //    // Remove a seção depois de ter excluído as webparts
      //    page.Sections.Remove(section);
      //  }
      //}
      ////page.RemovePageHeader();

      //page.AddSection(CanvasSectionTemplate.TwoColumnVerticalSection, 1);
      //ClientSideText txt1 = new ClientSideText() { Text = "testando" };
      //page.AddControl(txt1, page.Sections[0].Columns[0], 0);

      //page.Save();
      //page.Publish();
      //page.DisableComments();



      //bool append = false;
      //string fileName = "Components.csv";
      //string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName);
      //StringBuilder csv = new StringBuilder();
      //string finalField = "";
      //if (!System.IO.File.Exists(filePath))
      //{
      //  append = false;

      //}
      //else
      //{
      //  append = true;
      //}
      //StreamWriter writer = new StreamWriter(filePath, append);

      // Escreva o cabeçalho do arquivo CSV, separando os nomes dos campos por ";"
      //writer.WriteLine("Id;Title;Description;Propreties");
      //foreach (ClientSideComponent component in components)
      //{
      //  if (component.ComponentType == 1)
      //  {
      //    List<string> values = new List<string>();
      //    //values.Add(component.Manifest.ToString().Replace("\r", "").Replace("\n", "").Replace(";", "[.,]"));
      //    dynamic manifest = Newtonsoft.Json.Linq.JObject.Parse(component.Manifest.ToString());
      //    values.Add(manifest.id.Value);
      //    //if(manifest.preconfiguredEntries[0].group != null) values.Add(manifest.preconfiguredEntries[0].group["default"].Value);
      //    values.Add(manifest.preconfiguredEntries[0].title["default"].Value);
      //    values.Add(manifest.preconfiguredEntries[0].description["default"].Value);
      //    //values.Add(manifest.propertiesMetadata[0].current);



      //    ClientSideWebPart clientWp = new ClientSideWebPart(component) { Order = 1 };
      //    values.Add(clientWp.Properties.ToString().Replace("\r", "").Replace("\n", "").Replace(";", "[.,]"));
      //    string line = string.Join(";", values);
      //    writer.WriteLine(line);
      //  }

      //}
      //writer.Close();



      //Criar um novo link 
      //NavigationNodeCreationInformation navigationNodeCreationInformation = new NavigationNodeCreationInformation();
      //navigationNodeCreationInformation.Title = "Abre em nova janela";
      //navigationNodeCreationInformation.Url = "https://www.codesharepoint.com/csom/create-quick-launch-link-in-sharepoint-using-csom";
      //navigationNodeCreationInformation.AsLastNode = true;
      //quickLaunchCollection.Add(navigationNodeCreationInformation);
      //context.Load(quickLaunchCollection);
      //context.ExecuteQuery();

      //editar um link
      //Microsoft.SharePoint.Client.NavigationNode navigationNode = quickLaunchCollection.Where(n => n.Title == "New navigation").FirstOrDefault();
      //navigationNode.Title = "Editado via CSOM";
      //navigationNode.Url = "https://www.sharepointpals.com";

      //Microsoft.SharePoint.Client.NavigationNodeCollection quickLaunchCollection = web.Navigation.QuickLaunch;
      //context.Load(quickLaunchCollection);
      //context.ExecuteQueryRetry();

      //List<Microsoft.SharePoint.Client.NavigationNode> nodes = new List<Microsoft.SharePoint.Client.NavigationNode>();
      //foreach(Microsoft.SharePoint.Client.NavigationNode navigationNode in quickLaunchCollection)
      //{
      //  nodes.Add(navigationNode);
      //}


      //foreach (Microsoft.SharePoint.Client.NavigationNode newnodes in nodes)
      //{
      //  Microsoft.SharePoint.Client.NavigationNode navigationNode = quickLaunchCollection.Where(n => n.Title == newnodes.Title).FirstOrDefault();
      //  navigationNode.DeleteObject();

      //}
      //context.ExecuteQuery();

      //Microsoft.SharePoint.Client.NavigationNode rnode = nodes.Where(n => n.Title == "Abre em nova janela").FirstOrDefault();
      //nodes.Remove(rnode);
      //nodes.Insert(0, rnode);
      //foreach (Microsoft.SharePoint.Client.NavigationNode newnodes in nodes)
      //{
      //  NavigationNodeCreationInformation navigationNodeCreationInformation = new NavigationNodeCreationInformation();
      //  navigationNodeCreationInformation.Title = newnodes.Title;
      //  navigationNodeCreationInformation.Url = newnodes.Url;
      //  navigationNodeCreationInformation.AsLastNode = true;
      //  quickLaunchCollection.Add(navigationNodeCreationInformation);
      //}
      //context.Load(quickLaunchCollection);
      //context.ExecuteQuery();

      //criar nó filho 

      //Microsoft.SharePoint.Client.NavigationNode parentNode = quickLaunchCollection.Where(n => n.Title == "New navigation").FirstOrDefault();
      //NavigationNodeCreationInformation nodeCreationInformation = new NavigationNodeCreationInformation();

      //nodeCreationInformation.Title = "Filho new navigation";
      //nodeCreationInformation.Url = "https://google.com";

      //parentNode.Children.Add(nodeCreationInformation);
      //parentNode.Update();

      //context.ExecuteQueryRetry();

      // deletar link 
      //Microsoft.SharePoint.Client.NavigationNode navigationNode = quickLaunchCollection.Where(n => n.Title == "Editado via CSOM").FirstOrDefault();
      //navigationNode.DeleteObject();
      //context.ExecuteQuery();








      //web.ApplyTheme("Branco Queimado");

      //// Obter informações sobre os temas disponíveis
      //ThemeInfoCollection themes = context.Web.ThemeInfo;
      //context.Load(themes);
      //context.ExecuteQuery();

      //// Encontrar o tema desejado pelo nome
      //string themeName = "NomeDoTema";
      //ThemeInfo theme = themes.FirstOrDefault(t => t.Name == themeName);

      //if (theme != null)
      //{
      //  // Aplicar o tema ao site
      //  context.Web.ApplyTheme(theme.ThemeJson, theme.FontSchemeUrl, theme.BackgroundImageUrl, true);
      //  context.Web.Update();
      //  context.ExecuteQuery();
      //  Console.WriteLine("Tema aplicado com sucesso.");
      //}
      //else
      //{
      //  Console.WriteLine("O tema especificado não foi encontrado.");
      //}

      //Acessar a lista de temas, pegar pelo titulo, pegar as infos, e setar usanso o ApplyTheme.usar este link de referencia: https://medium.com/@softreeconsultings/how-to-change-the-look-of-sharepoint-using-csom-efc353e32516

      ////seta um tema
      //ThemeManager.ApplyTheme(web, "{\r\n  \"themePrimary\": \"#5eff00\",\r\n  \"themeLighterAlt\": \"#f9fff5\",\r\n  \"themeLighter\": \"#e5ffd6\",\r\n  \"themeLight\": \"#cfffb3\",\r\n  \"themeTertiary\": \"#9eff66\",\r\n  \"themeSecondary\": \"#71ff1f\",\r\n  \"themeDarkAlt\": \"#54e600\",\r\n  \"themeDark\": \"#47c200\",\r\n  \"themeDarker\": \"#348f00\",\r\n  \"neutralLighterAlt\": \"#1c1c1c\",\r\n  \"neutralLighter\": \"#252525\",\r\n  \"neutralLight\": \"#343434\",\r\n  \"neutralQuaternaryAlt\": \"#3d3d3d\",\r\n  \"neutralQuaternary\": \"#454545\",\r\n  \"neutralTertiaryAlt\": \"#656565\",\r\n  \"neutralTertiary\": \"#ffd9b3\",\r\n  \"neutralSecondary\": \"#ffb366\",\r\n  \"neutralPrimaryAlt\": \"#ff8f1f\",\r\n  \"neutralPrimary\": \"#ff8000\",\r\n  \"neutralDark\": \"#c26100\",\r\n  \"black\": \"#8f4700\",\r\n  \"white\": \"#121212\"\r\n}", "meu tema");
      ////muda as propriedades layout, visibilidade do titulo do site, tela de fundo e miniatura.
      //web.HeaderLayout = HeaderLayoutType.Extended;
      //web.HideTitleInHeader = true;
      //web.HeaderEmphasis = SPVariantThemeType.Neutral;



      //context.Load(web);
      //context.Load(web.AllProperties);

      //context.ExecuteQuery();
      //web.SiteLogoUrl = "https://classsolutions.sharepoint.com/sites/TestesModerno/SiteAssets/logo%20(2).png";
      //web.LogoAlignment = LogoAlignment.Left;
      //web.AllProperties["RectSiteLogoUrl"] = "https://classsolutions.sharepoint.com/sites/TestesModerno/SiteAssets/logo%20(2).png";



      //context.ExecuteQuery();
      ////Configurações de navegação
      //web.QuickLaunchEnabled = false;
      //web.MegaMenuEnabled = false;

      ////Configurações de rodapé
      //web.FooterEnabled = true;
      //web.FooterLayout = FooterLayoutType.Extended;
      //web.FooterEmphasis = FooterVariantThemeType.Strong;
      //web.SetFooterTitle("Footer");

      //web.SetFooterLogoUrl("https://classsolutions.sharepoint.com/sites/TestesModerno/SiteAssets/logo%20(2).png");

      //web.Update();
      //context.ExecuteQuery();

      //var list = context.Web.GetList("http://equipes/sites/RelatoriosEstudos/Lists/EncerramentoBradesco/SolicitarBaixaGarantia.aspx");
      //context.Load(list.Author);
      //context.ExecuteQuery();
      //Console.WriteLine(list.Author);
      //Console.ReadLine();


      // 

      //web.IsMultilingual = false;
      //web.AddSupportedUILanguage(1033); // English
      //web.AddSupportedUILanguage(1031); // German
      //web.AddSupportedUILanguage(1036); // French
      //web.AddSupportedUILanguage(1046); // Portugese (Brazil)
      //web.AddSupportedUILanguage(1049); // Russian
      //web.AddSupportedUILanguage(2052); // Chinese (Simplified)
      //web.AddSupportedUILanguage(3082); // Spanish
      //web.Update();
      //context.ExecuteQuery();

      //Microsoft.SharePoint.Client.RegionalSettings regionalSettings = web.RegionalSettings;
      //context.Load(web, w => w.RegionalSettings, w => w.RegionalSettings.TimeZones);
      //context.ExecuteQuery();

      //// Note: The Locale ID controls the numbering, sorting, calendar, and time formatting for the Web site
      //// 1033 is for English (United States)
      //web.RegionalSettings.LocaleId = uint.Parse("1033");

      //// Set Work Start Time
      //web.RegionalSettings.WorkDayStartHour = 8;

      //// Set Work End Time
      //web.RegionalSettings.WorkDayEndHour = 5;

      //// Work days: 127 = All, 0 = None, 64 = Sunday, 32 = Monday, 16 = Tuesday
      //// 8 = Wednesday, 4 = Thursday, 2 = Friday, 1 = Sunday, 62 = Monday to Friday
      //web.RegionalSettings.WorkDays = 62;

      //// Specify whether you want to use 12-hour time format or 24-hour format.
      //web.RegionalSettings.Time24 = true;

      //// Specify a secondary calendar that provides extra information on the calendar features
      //// 1 = Gregorain(Localized), 2 = Gregorian (always English strings), 3 = Japanese Era: Year of the Emperor
      //// 4 =        Year of Taiwan, 5 = Tangun Era (Korea),6 = Hijri, 7 = Thai
      //web.RegionalSettings.AlternateCalendarType = 1;

      //// Specify the type of calendar
      //// 0 — None, 1 — Gregorian, 3 — Japanese Emperor Era, 5 — Korean Tangun Era
      //// 6 — Hijri, 7 — Buddhist, 8 — Hebrew Lunar,  9 — Gregorian Middle East French Calendar
      //// 10 — Gregorian Arabic Calendar, 11 — Gregorian Transliterated English Calendar
      //// 12 — Gregorian Transliterated French Calendar, 16 — Saka Era
      //web.RegionalSettings.CalendarType = 1;

      //// Set first day of week
      //// 0 = Sunday and 6 = Saturday(last day)
      //web.RegionalSettings.FirstDayOfWeek = 1;

      //// Set first week of year
      //// 0 = Starts on Jan 1, 1 = First full week, 2 = First 4-day week
      //web.RegionalSettings.FirstWeekOfYear = 1;

      //// Show week numbers in the Date Navigator
      //web.RegionalSettings.ShowWeeks = true;

      //Microsoft.SharePoint.Client.TimeZone utcTimeZone = web.RegionalSettings.TimeZones.Where(timezone => timezone.Id == 8).First();


      //web.RegionalSettings.TimeZone = utcTimeZone;

      //web.RegionalSettings.Update();

      //context.ExecuteQuery();
       JObject GetQuickLinkItem(int quickLinkItemId, string siteUrl, string siteId, string webId, QuickLinkItem quickLinkItem)
      {
        var siteUri = new Uri(siteUrl);
        bool external = quickLinkItem.Url[0] != '/' && !quickLinkItem.Url.StartsWith($"https://{siteUri.Host}", StringComparison.CurrentCultureIgnoreCase);
        string blankGuid = new Guid().ToString();

        JObject item = JObject.FromObject(new
        {
          id = quickLinkItemId,
          itemType = external ? 2 : 5,
          siteId = external ? blankGuid : siteId,
          webId = external ? blankGuid : webId,
          uniqueId = quickLinkItem.UniqueId,
          fileExtension = quickLinkItem.Url,
          progId = "",
          activity = new
          {
            activity = "Added a few seconds ago",
            people = new object[]
                {
                new
                {
                    name = "SharePoint",
                    profileImageSrc = "https://spoprod-a.akamaihd.net/files/sp-client-prod_2017-09-29.011/icon_sp-icon_fd9ed950.png",
                    initials = ""
                }
                }
          },
          hasInvalidUrl = !external,
          flags = 0,
          renderInfo = quickLinkItem.RenderInfo
        });

        return item;
      }

       JObject GetQuickLinksServerProcessedContent(string webUrl, string siteId, string webId, List<QuickLinkItem> quickLinkItems)
      {
        JObject searchablePlainTexts = new JObject();
        JObject imageSources = new JObject();
        JObject links = new JObject();

        for (var index = 0; index < quickLinkItems.Count; index++)
        {
          var quickLink = quickLinkItems[index];
          var renderInfoLinkUrl = GetRenderInfoLinkUrl(webUrl, siteId, webId, quickLink.UniqueId);
          searchablePlainTexts.Add($"items[{index}].title", quickLink.Title);
          searchablePlainTexts.Add($"items[{index}].caption", string.Empty);
          imageSources.Add($"items[{index}].pictureUrl", string.Empty);
          links.Add($"items[{index}].url", quickLink.Url);
          links.Add($"items[{index}].renderInfo.linkUrl", renderInfoLinkUrl);
        }

        JObject serverProcessedContent = JObject.FromObject(new
        {
          htmlStrings = new JObject(),
          searchablePlainTexts = searchablePlainTexts,
          imageSources = imageSources,
          links = links
        });

        return serverProcessedContent;
      }

      string GetRenderInfoLinkUrl(string webUrl, string siteId, string webId, string uniqueId)
      {
        return $"{webUrl}/_layouts/15/ShortcutLink.aspx?siteId={siteId}&webId={webId}&uniqueId={uniqueId}&web=1";
      }

       ClientSideWebPart CreateQuickLinksWebPart(ClientContext clientContext, ClientSidePage page, List<QuickLinkItem> quickLinks)
      {
        clientContext.Load(clientContext.Site, s => s.Id);
        clientContext.Load(clientContext.Web, w => w.Id, w => w.Url);
        clientContext.Load(page.PageListItem, p => p.Id);
        clientContext.ExecuteQueryRetry();

        string siteId = clientContext.Site.Id.ToString();
        string webId = clientContext.Web.Id.ToString();

        JObject serverProcessedContent = GetQuickLinksServerProcessedContent(clientContext.Web.Url, siteId, webId, quickLinks);

        var quickLinksWebPart = page.InstantiateDefaultWebPart(DefaultClientSideWebParts.QuickLinks);
        quickLinksWebPart.PropertiesJson = JObject.FromObject(new
        {
          controlType = 3,
          displayMode = 2,
          id = "05924478-460a-417a-afa8-83792656eff7",
          position = new
          {
            zoneIndex = 1,
            sectionIndex = 1,
            controlIndex = 1
          },
          webPartId = quickLinksWebPart.WebPartId,
          webPartData = new
          {
            id = quickLinksWebPart.WebPartId,
            instanceId = Guid.NewGuid().ToString(),
            title = "Quick links",
            description = "Add links to important documents and pages.",
            serverProcessedContent = serverProcessedContent,
            dataVersion = "1.0",
            properties = new
            {
              items = quickLinks.Select((QuickLinkItem quickLink, int index) => GetQuickLinkItem(index + 1, clientContext.Web.Url, siteId, webId, quickLink)),
              isMigrated = true,
              layoutId = "FilmStrip",
              shouldShowThumbnail = true,
              hideWebPartWhenEmpty = true,
              dataProviderId = "QuickLinks",
              webId = webId,
              siteId = siteId,
              baseUrl = clientContext.Web.Url
            }
          }
        }).ToString();

        return quickLinksWebPart;
      }





    }
  }
}
