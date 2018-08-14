using Microsoft.Owin;
using Microsoft.Owin.Security.DataProtection;
using Owin;

[assembly: OwinStartupAttribute(typeof(StudyGuide.Startup))]
namespace StudyGuide
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            UnityConfig.RegisterComponents(app.GetDataProtectionProvider());
            ConfigureAuth(app);
        }
    }
}
