using System.Diagnostics;
using GameZone.Models;
using GameZone.Services;
using Microsoft.AspNetCore.Mvc;

namespace GameZone.Controllers
{
    public class HomeController : Controller
    {
        private readonly IGameServices gameServices;

        public HomeController(IGameServices gameServices)
        {
            this.gameServices = gameServices;
        }

        public IActionResult Index()
        {
            var games = gameServices.GetAll();
            return View(games);
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
