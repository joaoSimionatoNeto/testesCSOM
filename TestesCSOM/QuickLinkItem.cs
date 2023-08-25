using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static TestesCSOM.Program;

namespace TestesCSOM
{
  public class QuickLinkItem
  {
    public string UniqueId { get; set; }
    public string Url { get; set; }
    public string Title { get; set; }
    public QuickLinkRenderInfo RenderInfo { get; set; }
  }
}
