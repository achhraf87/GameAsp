using Microsoft.AspNetCore.Mvc;
using GameZone.ViewModels;
using GameZone.Data;
using Microsoft.AspNetCore.Mvc.Rendering;
using GameZone.Services;
using GameZone.Models;
using Microsoft.EntityFrameworkCore;

namespace GameZone.Controllers
{
    public class GameController : Controller
    {
        private readonly ICategoriesService _categoriesService;
        private readonly IDevicesService _devicesService;
        private readonly IGameServices _gamesService;

        public GameController(ICategoriesService categoriesService,
            IDevicesService devicesService,
            IGameServices gamesService)
        {
            _categoriesService = categoriesService;
            _devicesService = devicesService;
            _gamesService = gamesService;
        }

        //public IActionResult Index(int pageNumber = 1, int pageSize = 5)
        //{
        //    var games = _gamesService.GetAll().Skip((pageNumber - 1) * pageSize).Take(pageSize).ToList();
        //    var totalGames = _gamesService.GetAll().Count();
        //    var totalPages = (int)Math.Ceiling(totalGames / (double)pageSize);

        //    var paginationInfo = new PaginationInfo
        //    {
        //        PageNumber = pageNumber,
        //        TotalPages = totalPages
        //    };

        //    ViewData["PaginationInfo"] = paginationInfo;

        //    return View(games);
        //}
        public IActionResult Index(string searchTerm, int pageNumber = 1, int pageSize = 5)
        {
            var query = _gamesService.GetAll();

            if (!string.IsNullOrEmpty(searchTerm))
            {
                query = query.Where(g => g.Name.Contains(searchTerm) || g.Category.Name.Contains(searchTerm));
            }

            var games = query.Skip((pageNumber - 1) * pageSize).Take(pageSize).ToList();
            var totalGames = query.Count();
            var totalPages = (int)Math.Ceiling(totalGames / (double)pageSize);

            var paginationInfo = new PaginationInfo
            {
                PageNumber = pageNumber,
                TotalPages = totalPages
            };

            ViewData["PaginationInfo"] = paginationInfo;

            return View(games);
        }




        //public IActionResult Index()
        //{
        //    var games = _gamesService.GetAll();
        //    return View(games);
        //}

        public IActionResult Details(int id)
        {
            var game = _gamesService.GetById(id);

            if (game is null)
                return NotFound();

            return View(game);
        }

        [HttpGet]
        public IActionResult Create()
        {
            CreateGameViewModel viewModel = new()
            {
                Categories = _categoriesService.GetListItems(),
                Devices = _devicesService.GetListItems()
            };

            return View(viewModel);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(CreateGameViewModel model)
        {
            if (!ModelState.IsValid)
            {
                model.Categories = _categoriesService.GetListItems();
                model.Devices = _devicesService.GetListItems();
                return View(model);
            }

            await _gamesService.Create(model);

            return RedirectToAction(nameof(Index));
        }

        [HttpGet]
        public IActionResult Edit(int id)
        {
            var game = _gamesService.GetById(id);

            if (game is null)
                return NotFound();

            EditGameViewModel viewModel = new()
            {
                Id = id,
                Name = game.Name,
                Description = game.Description,
                CategoryId = game.CategoryId,
                SelectedDevices = game.GameDevices.Select(d => d.DeviceId).ToList(),
                Categories = _categoriesService.GetListItems(),
                Devices = _devicesService.GetListItems(),
                CurrentCover = game.Cover
            };

            return View(viewModel);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(EditGameViewModel model)
        {
            if (!ModelState.IsValid)
            {
                model.Categories = _categoriesService.GetListItems();
                model.Devices = _devicesService.GetListItems();
                return View(model);
            }

            var game = await _gamesService.Update(model);

            if (game is null)
                return BadRequest();

            return RedirectToAction(nameof(Index));
        }

        [HttpDelete]
        public IActionResult Delete(int id)
        {
            var isDeleted = _gamesService.Delete(id);

            return isDeleted ? Ok() : BadRequest();
        }
    }
}
