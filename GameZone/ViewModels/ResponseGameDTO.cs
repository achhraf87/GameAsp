using GameZone.Models;

namespace GameZone.ViewModels
{
    public class ResponseGameDTO
    {
        public List<Game> Data { get; set; }
        public int RecordFiltred  { get; set; }
        public int RecordTotal { get; set; }
        public PaginationInfo pagination { get; set; } 
    }
}
