using Microsoft.AspNetCore.Mvc;
using GameZone.ViewModels;
using GameZone.Data;
using Microsoft.AspNetCore.Mvc.Rendering;
using GameZone.Services;
using GameZone.Models;
using Microsoft.EntityFrameworkCore;
using System.Composition;
using ClosedXML.Excel;

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
            ViewBag.SearchTerm = searchTerm;

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

        public async Task<IActionResult> ExportToExcel(string searchTerm)
        {
            var query = _gamesService.GetAll(); // Remplacez cela par votre méthode pour récupérer tous les jeux depuis le service.

            // Appliquer le filtre de recherche si un terme de recherche est spécifié
            if (!string.IsNullOrEmpty(searchTerm))
            {
                query = query.Where(g => g.Name.Contains(searchTerm) || g.Category.Name.Contains(searchTerm));
            }

            var games = query.ToList();
           

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Games");
                var currentRow = 1;

                // Ajouter les en-têtes
                worksheet.Cell(currentRow, 1).Value = "Name";
                worksheet.Cell(currentRow, 2).Value = "Category";

                // Remplir les données
                foreach (var game in games)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = game.Name;
                    worksheet.Cell(currentRow, 2).Value = game.Category.Name;
                }

                // Rendre le fichier de travail en mémoire
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    // Retourner le fichier Excel en tant que résultat de l'action
                    return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "les candidats.xlsx");
                }
            }
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
