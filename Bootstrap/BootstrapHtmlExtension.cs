using System.Web.Mvc;
using Bootstrap.BootstrapMethods;

namespace Bootstrap
{
    public static class BootstrapHtmlExtension
    {
        public static Bootstrap<TModel> Bootstrap<TModel>(this HtmlHelper<TModel> helper)
        {
            return new Bootstrap<TModel>(helper);
        }
    }
}
