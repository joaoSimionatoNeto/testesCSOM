using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static TestesCSOM.Program;

namespace TestesCSOM
{
  public class QuickLinkRenderInfo
  {
    public string imageUrl { get; set; }
    public QuickLinkCompactImageInfo compactImageInfo { get; set; }
    public string backupImageurl { get; set; }
    public string iconUrl { get; set; }
    public string accentColor { get; set; }
    public int imageFit { get; set; }
    public bool forceStandardImageSize { get; set; }

    public static QuickLinkRenderInfo DocLibrary = new QuickLinkRenderInfo
    (
        "https://spoprod-a.akamaihd.net/files/sp-client-prod_2017-09-29.016/icon_doc-library_04a9a5ce.png",
        "https://spoprod-a.akamaihd.net/files/sp-client-prod_2017-09-29.016/icon_genericfile_56b72b7b.png",
        "DocLibrary"
    );

    public static QuickLinkRenderInfo Globe = new QuickLinkRenderInfo
    (
        "https://spoprod-a.akamaihd.net/files/sp-client-prod_2017-09-29.016/icon_link_22d1503c.png",
        "https://spoprod-a.akamaihd.net/files/sp-client-prod_2017-09-29.016/icon_genericfile_56b72b7b.png",
        "Globe"
    );

    public QuickLinkRenderInfo(string imageUrl, string iconUrl, string iconName)
    {
      this.imageUrl = imageUrl;
      compactImageInfo = new QuickLinkCompactImageInfo
      {
        iconName = iconName,
        color = string.Empty,
        imageUrl = imageUrl,
        forceIconSize = true
      };
      backupImageurl = string.Empty;
      this.iconUrl = iconUrl;
      accentColor = string.Empty;
      imageFit = 0;
      forceStandardImageSize = true;
      isFetching = false;
    }
  }
}
