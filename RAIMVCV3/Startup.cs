using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(RAIMVCV3.Startup))]
namespace RAIMVCV3
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
