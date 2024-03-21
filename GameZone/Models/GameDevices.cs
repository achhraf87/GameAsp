using System.ComponentModel.DataAnnotations.Schema;

namespace GameZone.Models
{
    public class GameDevices
    {
        [ForeignKey(nameof(Game))]
        public int GameId { get; set; }
        public Game Game { get; set; }

        [ForeignKey(nameof(Device))]
        public int DeviceId { get; set; }
        public Device Device { get; set; }
    }
}
