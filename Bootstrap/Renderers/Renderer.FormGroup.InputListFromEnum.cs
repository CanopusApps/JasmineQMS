using System.Web.Mvc;
using Bootstrap.ControlModels;
using Bootstrap.Controls;
using Bootstrap.TypeExtensions;

namespace Bootstrap.Renderers
{
    internal static partial class Renderer
    {
        public static string RenderFormGroupInputListFromEnum(HtmlHelper html, BootstrapInputListFromEnumModel model)
        {
            var input = Renderer.RenderInputListFromEnum(html, model);

            string label = Renderer.RenderLabel(html, new BootstrapLabelModel
            {
                htmlFieldName = model.htmlFieldName,
                metadata = model.metadata,
                htmlAttributes = new { @class = "control-label" }.ToDictionary()
            });

            bool fieldIsValid = true;
            if(model != null) fieldIsValid = html.ViewData.ModelState.IsValidField(model.htmlFieldName);
            return new BootstrapFormGroup(input, label, FormGroupType.textboxLike, fieldIsValid).ToHtmlString();
        }
    }
}
