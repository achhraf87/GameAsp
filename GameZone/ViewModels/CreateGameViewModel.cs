
using GameZone.Attributes;
using GameZone.Settings;


namespace GameZone.ViewModels
{
	public class CreateGameViewModel: GameViewModel
	{

		[AllowedExtenstion(FileSetting.imageExtension)]
		[AllowedFileSize(FileSetting.maxFileSizeInB)]
		public IFormFile Cover { get; set; } = default!;
	}
}
