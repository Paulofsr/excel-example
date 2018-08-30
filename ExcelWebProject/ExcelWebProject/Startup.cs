using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(ExcelWebProject.Startup))]
namespace ExcelWebProject
{
    public partial class Startup {
        public void Configuration(IAppBuilder app) {
            ConfigureAuth(app);
        }
    }
}
