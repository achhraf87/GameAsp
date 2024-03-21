using Microsoft.AspNetCore.Mvc.Rendering;

namespace GameZone.Services
{
    public interface ICategoriesService
    {
        public IEnumerable<SelectListItem> GetListItems();
    }
}
