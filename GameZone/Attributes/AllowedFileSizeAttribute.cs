using System.ComponentModel.DataAnnotations;
using GameZone.Settings;

namespace GameZone.Attributes
{
    public class AllowedFileSizeAttribute: ValidationAttribute
    {
        private readonly int _allowedFileSize;

        public AllowedFileSizeAttribute(int allowedFileSize)
        {
            _allowedFileSize = allowedFileSize;
        }

        protected override ValidationResult? IsValid(object? value, ValidationContext validationContext)
        {
            var file = value as IFormFile;
            if (file is not null)
            {
                var size = file.Length;
                if (size > _allowedFileSize)
                {
                    return new ValidationResult($"Size must be less than {FileSetting.maxFileSizeInMB} MB");
                }
            }
            return ValidationResult.Success;
        }
    }
}
