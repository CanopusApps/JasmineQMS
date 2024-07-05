using System.Collections.Generic;
using Bootstrap.Infrastructure;

namespace Bootstrap
{
    public class FormActions : HtmlElement
    {
        public FormActions()
            : base("div")
        {
            EnsureClass("form-actions");
        }

        public FormActions HtmlAttributes(IDictionary<string, object> htmlAttributes)
        {
            SetHtmlAttributes(htmlAttributes);
            return this;
        }

        public FormActions HtmlAttributes(object htmlAttributes)
        {
            SetHtmlAttributes(htmlAttributes);
            return this;
        }
    }
}
