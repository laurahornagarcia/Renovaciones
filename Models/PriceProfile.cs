using System;
using System.Collections.Generic;

namespace Ofertum.Renovaciones.Models
{
    public class PriceProfile
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public string? Name { get; set; }
        public Dictionary<string, string>? Prices { get; set; }
        public DateTime UpdatedAtUtc { get; set; } = DateTime.UtcNow;
    }
}
