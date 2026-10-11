using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using WordServer.Api.Data;

namespace WordServer.Api.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class HealthController : ControllerBase
    {
        private readonly AppDbContext _db;

        public HealthController(AppDbContext db)
        {
            _db = db;
        }

        [HttpGet]
        public async Task<IActionResult> GetHealth()
        {
            bool dbConnected = false;
            string dbProvider = "Unknown";
            try
            {
                dbConnected = await _db.Database.CanConnectAsync();
                dbProvider = _db.Database.ProviderName ?? "Unknown";
            }
            catch
            {
                dbConnected = false;
            }

            return Ok(new
            {
                status = dbConnected ? "Healthy" : "Degraded",
                databaseConnected = dbConnected,
                databaseProvider = dbProvider,
                timestamp = DateTime.UtcNow,
                version = "1.0.0"
            });
        }
    }
}
