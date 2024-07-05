using System.Web.Mvc;
using Bootstrap.Infrastructure;

namespace Bootstrap.Controls
{
    public class TableRowBuilder<TModel> : BuilderBase<TModel, TableRow>
    {
        internal TableRowBuilder(HtmlHelper<TModel> htmlHelper, TableRow tableRow)
            : base(htmlHelper, tableRow)
        {
        }
    }
}
