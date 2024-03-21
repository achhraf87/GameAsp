using System.ComponentModel.DataAnnotations;
using GameZone.Settings;

namespace GameZone.Attributes
{
    public class AllowedExtenstionAttribute: ValidationAttribute
    {
        private readonly string _allowedExtension;
        public AllowedExtenstionAttribute(string allowedExtension)
        {
            _allowedExtension = allowedExtension;
        }
        protected override ValidationResult? IsValid(object? value, ValidationContext validationContext)
        {
            var file = value as IFormFile;
            if (file is not null)
            {
                var allowed = _allowedExtension.Split(",").Contains(Path.GetExtension(file.FileName));
                if (!allowed)
                {
                    return new ValidationResult($"Only {_allowedExtension} are allowed!");
                }

            }
            return ValidationResult.Success;
        }
    }
}
