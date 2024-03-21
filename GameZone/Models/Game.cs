using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace GameZone.Models
{
    public class Game: BaseEntity
    {
        [MaxLength(2500)]
        public string Description { get; set; }
        [MaxLength(500)]
        public string Cover { get; set; }

        [ForeignKey(nameof(Category))]
        public int CategoryId { get; set; }

        public Category Category { get; set; }

        public ICollection<GameDevices> GameDevices { get; set; } = new List<GameDevices>();
    }
}
