using Microsoft.Owin;
using Owin;

[assembly: OwinStartup(typeof(MVCWebSite.Startup))]
namespace MVCWebSite
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
