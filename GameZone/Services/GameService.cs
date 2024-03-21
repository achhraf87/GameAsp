using System.Drawing;
using GameZone.Data;
using GameZone.Models;
using GameZone.Settings;
using GameZone.ViewModels;
using Microsoft.EntityFrameworkCore;

namespace GameZone.Services
{
    public class GameService : IGameServices
    {
        private readonly ApplicationDbContext _context;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private string _pathName;

        public GameService(ApplicationDbContext context, IWebHostEnvironment webHostEnvironment)
        {
            _context = context;
            _webHostEnvironment = webHostEnvironment;
            _pathName = $"{_webHostEnvironment.WebRootPath}{FileSetting.fileName}";
        }

        public async Task Create(CreateGameViewModel model)
        {
            var imageName = $"{Guid.NewGuid()}{Path.GetExtension(model.Cover.FileName)}";
            var path = Path.Combine(_pathName, imageName);
            using var stream = File.Create(path);
            await model.Cover.CopyToAsync(stream);

            Game game = new()
            {
                Name = model.Name,
                Description = model.Description,
                CategoryId = model.CategoryId,
                Cover = imageName,
                GameDevices = model.SelectedDevices.Select(d => new GameDevices { DeviceId = d }).ToList(),
            };

            _context.Add(game);
            _context.SaveChanges();
        }

        public IEnumerable<Game> GetAll()
        {
            return _context
                .Games
                .Include(g => g.Category)
                .Include(g => g.GameDevices)
                .ThenInclude(gd => gd.Device)
                .AsNoTracking()
                .ToList();
        }

        public Game? GetById(int id)
        {
            return _context
                .Games
                .Include(g => g.Category)
                .Include(g => g.GameDevices)
                .ThenInclude(d => d.Device)
                .AsNoTracking()
                .SingleOrDefault(g => g.Id == id);
        }

        public async Task<Game?> Update(EditGameViewModel model)
        {
            var game = _context.Games
                .Include(g => g.GameDevices)
                .SingleOrDefault(g => g.Id == model.Id);

            if (game is null)
                return null;

            var hasNewCover = model.Cover is not null;
            var oldCover = game.Cover;

            game.Name = model.Name;
            game.Description = model.Description;
            game.CategoryId = model.CategoryId;
            game.GameDevices = model.SelectedDevices.Select(d => new GameDevices { DeviceId = d }).ToList();

            if (hasNewCover)
            {
                game.Cover = await SaveCover(model.Cover!);
            }

            var effectedRows = _context.SaveChanges();

            if (effectedRows > 0)
            {
                if (hasNewCover)
                {
                    var cover = Path.Combine(_pathName, oldCover);
                    File.Delete(cover);
                }

                return game;
            }
            else
            {
                var cover = Path.Combine(_pathName, game.Cover);
                File.Delete(cover);

                return null;
            }
        }

        public async Task<string?> SaveCover(IFormFile cover)
        {
            var newImage = $"{Guid.NewGuid()}{Path.GetExtension(cover.FileName)}";
            var newPath = Path.Combine(_pathName, newImage);

            using var stream = File.Create(newPath);
            await cover.CopyToAsync(stream);

            return newImage;

        }

        public bool Delete(int id)
        {
            var isDeleted = false;
            var game = _context.Games.SingleOrDefault(g => g.Id == id);
            
            if (game is null)
                return isDeleted;

            _context.Remove(game);
            var effectedRows = _context.SaveChanges();

            if(effectedRows > 0)
            {
                isDeleted = true;
                var cover = Path.Combine(_pathName, game.Cover);
                File.Delete(cover);
            }
            return isDeleted;
        }
    }
}
