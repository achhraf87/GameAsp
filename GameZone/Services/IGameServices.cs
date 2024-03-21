using GameZone.ViewModels;
using GameZone.Models;

namespace GameZone.Services
{
    public interface IGameServices
    {
        IEnumerable<Game> GetAll();
        Task Create(CreateGameViewModel model);
        Game? GetById(int id);

        Task<Game?> Update(EditGameViewModel game);

        bool Delete(int id);
    }
}
