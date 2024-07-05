using System.Collections.Generic;
using Bootstrap.Infrastructure;

namespace Bootstrap
{
    public class ButtonGroup : HtmlElement
    {
        public ButtonGroup()
            : base("div")
        {
            EnsureClass("btn-group");
        }

        public ButtonGroup HtmlAttributes(IDictionary<string, object> htmlAttributes)
        {
            SetHtmlAttributes(htmlAttributes);
            return this;
        }

        public ButtonGroup HtmlAttributes(object htmlAttributes)
        {
            SetHtmlAttributes(htmlAttributes);
            return this;
        }
    }
}
