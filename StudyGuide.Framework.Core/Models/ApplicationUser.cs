using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.EntityFramework;
using StudyGuide.Framework.Core.Interface;
using System;
using System.Security.Claims;
using System.Threading.Tasks;

namespace StudyGuide.Framework.Core.Models
{
    public class ApplicationUser : IdentityUser, ITimestamp
    {
        /// <summary>
        /// The full name of the user
        /// </summary>
        public string FullName { get; set; }

        /// <summary>
        /// If a user changes their email, this field will hold the new email until it has been verified
        /// </summary>
        public string UnverifiedEmailAddress { get; set; }

        //public ISet<Chat> Chats { get; set; }

        public async Task<ClaimsIdentity> GenerateUserIdentityAsync(UserManager<ApplicationUser> manager)//, string authenticationType)
        {
            //var userIdentity = await manager.CreateIdentityAsync(this, authenticationType);
            var userIdentity = await manager.CreateIdentityAsync(this, DefaultAuthenticationTypes.ApplicationCookie);

            //userIdentity.AddClaim(new Claim("tenantId", Tenant.TenantId.ToString()));
            return userIdentity;
        }
        
        public DateTimeOffset CreatedAt { get; set; }
        public DateTimeOffset UpdatedAt { get; set; }
    }
}