using System.Web;
using System.Web.Mvc;

namespace TEPLQMS
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
            //filters.Add(new CustomFilter());
        }
    }
    public class CustomFilter : FilterAttribute, IExceptionFilter
    {
        public void OnException(ExceptionContext filterContext)
        {
            throw new System.NotImplementedException();
        }
    }
}
