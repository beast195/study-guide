using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.EntityFramework;
using Microsoft.AspNet.Identity.Owin;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.DataHandler;
using Microsoft.Owin.Security.DataHandler.Encoder;
using Microsoft.Owin.Security.DataHandler.Serializer;
using Microsoft.Owin.Security.DataProtection;
using Microsoft.Practices.Unity;
using StudyGuide.Common.Unity;
using StudyGuide.Framework.Core.Models;
using System.Web;
using System.Web.Caching;
using System.Web.Mvc;
using Unity.Mvc5;
using MvcUnityDependencyResolver = Unity.Mvc5.UnityDependencyResolver;
//using WebApiUnityDependencyResolver = Unity.WebApi.UnityDependencyResolver;

namespace StudyGuide
{
    public static class UnityConfig
    {
        public static UnityContainer RegisterComponents(IDataProtectionProvider dataProtectionProvider)
        {
            var container = new UnityContainer();

            DiContainerHelper.RegisterComponents<HierarchicalLifetimeManager>(container);

            container.RegisterType<SignInManager<ApplicationUser, string>, ApplicationSignInManager>(new HierarchicalLifetimeManager());
            container.RegisterType<IUserStore<ApplicationUser>, UserStore<ApplicationUser>>(new HierarchicalLifetimeManager());
            container.RegisterType<UserManager<ApplicationUser>, ApplicationUserManager>(new HierarchicalLifetimeManager());
            container.RegisterType<Cache, Cache>(new InjectionFactory(o => HttpContext.Current.Cache));
            container.RegisterType<IAuthenticationManager>(new InjectionFactory(o => HttpContext.Current.GetOwinContext().Authentication));

            container.RegisterType<IdentityFactoryOptions<ApplicationUserManager>>(
                new InjectionFactory(o => new IdentityFactoryOptions<ApplicationUserManager>
                {
                    DataProtectionProvider = dataProtectionProvider
                }));

            container.RegisterType<ISecureDataFormat<AuthenticationTicket>, SecureDataFormat<AuthenticationTicket>>();
            container.RegisterType<IDataSerializer<AuthenticationTicket>, TicketSerializer>();
            container.RegisterType<ITextEncoder, Base64UrlTextEncoder>();
            container.RegisterType<IDataProtector>(new InjectionFactory(o => dataProtectionProvider.Create()));

            DependencyResolver.SetResolver(new MvcUnityDependencyResolver(container));
            //GlobalConfiguration.Configuration.DependencyResolver = new MvcUnityDependencyResolver(container);
            //DependencyResolver.SetResolver(new UnityDependencyResolver(container));
            return container;

            
        }
    }
}