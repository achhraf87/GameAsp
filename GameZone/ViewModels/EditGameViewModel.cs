using GameZone.Attributes;
using GameZone.Settings;

namespace GameZone.ViewModels
{
    public class EditGameViewModel: GameViewModel
    {
        public int Id { get; set; }
        
        public string? CurrentCover { get; set; }
        [AllowedExtenstion(FileSetting.imageExtension)]
        [AllowedFileSize(FileSetting.maxFileSizeInB)]
        public IFormFile? Cover { get; set; } = default!;
    }
}
