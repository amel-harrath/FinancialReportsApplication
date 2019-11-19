using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(TASEF.Startup))]
namespace TASEF
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
