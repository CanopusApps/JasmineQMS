using System.Web.Mvc;
using Bootstrap.Infrastructure;

namespace Bootstrap.Controls
{
    public class FormActionsBuilder<TModel> : BuilderBase<TModel, FormActions>
    {
        internal FormActionsBuilder(HtmlHelper<TModel> htmlHelper, FormActions formActions)
            : base(htmlHelper, formActions)
        {
        }
    }
}
