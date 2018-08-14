using Microsoft.Practices.Unity;
using StudyGuide.Framework.Core.Boundaries;
using StudyGuide.Framework.Core.Models;
using StudyGuide.Framework.Repositories;
using StudyGuide.Logic.Boundaries;
using StudyGuide.Logic.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudyGuide.Common.Unity
{
    public static class DiContainerHelper
    {
        public static void RegisterComponents<T>(UnityContainer container) where T : LifetimeManager, new()
        {
            #region Services


            //container.RegisterType<IUserService, UserService>(new T());
            container.RegisterType<IMnemonicService, MnemonicService>(new T());
            container.RegisterType<IBookService, BookService>(new T());
            //container.RegisterType<ICompetitionService, CompetitionService>(new T());
            //container.RegisterType<IIdentityMessageService, SmsService>("SmsService", new T());
            //container.RegisterType<IIdentityMessageService, EmailService>("EmailService", new T());
            //container.RegisterType<IChatService, ChatService>(new T());
            //container.RegisterType<IAppCampaignService, AppCampaignService>(new T());

            #endregion Services

            #region Repos


            container.RegisterType<IBaseRepository<ApplicationUser>, BaseRepository<ApplicationUser>>(new T());
            container.RegisterType<IBaseRepository<SentenceLibrary>, BaseRepository<SentenceLibrary>>(new T());
            container.RegisterType<IBaseRepository<Book>, BaseRepository<Book>>(new T());
            container.RegisterType<IBaseRepository<Chapter>, BaseRepository<Chapter>>(new T());
            //container.RegisterType<IBaseRepository<Chat>, BaseRepository<Chat>>(new T());


            #endregion Repos




        }
    }
}
