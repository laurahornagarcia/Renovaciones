using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using Ofertum.Renovaciones.Models;

namespace Ofertum.Renovaciones.Services
{
    public class PriceProfileStore
    {
        private const string STORAGE_DIRECTORY = "App_Data/price-profiles";
        private const string INDEX_FILE = "profiles.index.json";

        private readonly JsonSerializerOptions _jsonSerializerOptions = new JsonSerializerOptions(JsonSerializerDefaults.Web)
        {
            WriteIndented = true
        };

        public PriceProfileStore()
        {
            _EnsureStorageCreated();
        }

        private void _EnsureStorageCreated()
        {
            Directory.CreateDirectory(STORAGE_DIRECTORY);
        }

        private async Task<List<string>> _ReadIndexAsync()
        {
            var indexPath = Path.Combine(STORAGE_DIRECTORY, INDEX_FILE);
            if (!File.Exists(indexPath))
            {
                return new List<string>();
            }

            var json = await File.ReadAllTextAsync(indexPath);
            return JsonSerializer.Deserialize<List<string>>(json, _jsonSerializerOptions) ?? new List<string>();
        }

        private async Task _WriteIndexAsync(List<string> index)
        {
            var indexPath = Path.Combine(STORAGE_DIRECTORY, INDEX_FILE);
            var json = JsonSerializer.Serialize(index, _jsonSerializerOptions);
            await File.WriteAllTextAsync(indexPath, json);
        }

        public async Task<List<PriceProfile>> ListAsync()
        {
            var ids = await _ReadIndexAsync();
            var profiles = new List<PriceProfile>();
            foreach (var id in ids)
            {
                var profile = await GetAsync(id);
                if (profile != null)
                {
                    profiles.Add(profile);
                }
            }
            return profiles;
        }

        public async Task<PriceProfile?> GetAsync(string id)
        {
            var filePath = Path.Combine(STORAGE_DIRECTORY, $"{id}.json");
            if (!File.Exists(filePath))
            {
                return null;
            }

            var json = await File.ReadAllTextAsync(filePath);
            return JsonSerializer.Deserialize<PriceProfile>(json, _jsonSerializerOptions);
        }

        public async Task SaveAsync(PriceProfile profile)
        {
            var filePath = Path.Combine(STORAGE_DIRECTORY, $"{profile.Id}.json");
            var json = JsonSerializer.Serialize(profile, _jsonSerializerOptions);
            await File.WriteAllTextAsync(filePath, json);

            var index = await _ReadIndexAsync();
            if (!index.Contains(profile.Id))
            {
                index.Add(profile.Id);
                await _WriteIndexAsync(index);
            }
        }

        public async Task<PriceProfile?> CloneAsync(string sourceId, string newName)
        {
            var sourceProfile = await GetAsync(sourceId);
            if (sourceProfile == null)
            {
                return null;
            }

            var newProfile = new PriceProfile
            {
                Name = newName,
                Prices = sourceProfile.Prices != null ? new Dictionary<string, string>(sourceProfile.Prices) : new Dictionary<string, string>()
            };

            await SaveAsync(newProfile);
            return newProfile;
        }

        public async Task DeleteAsync(string id)
        {
            var filePath = Path.Combine(STORAGE_DIRECTORY, $"{id}.json");
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            var index = await _ReadIndexAsync();
            if (index.Contains(id))
            {
                index.Remove(id);
                await _WriteIndexAsync(index);
            }
        }
    }
}
