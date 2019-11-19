using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Security.Principal;
using System.Web;

namespace TASEF.Helpers
{
    public static class IdentityExtensions
    {
        public static string GetDisplayName(this IIdentity identity)
        {
            if (identity == null)
            {
                throw new ArgumentNullException("identity");
            }
            var ci = identity as ClaimsIdentity;
            if (ci != null)
            {
                return ci.FindFirst("FirstName").ToString().Substring(("FirstName: ").Length); ;
            }
            return null;
        }
    }
}