using System.Text.Json;
using ZipStation.Worker.Entities;
using ZipStation.Worker.Helpers;
using ZipStation.Worker.Repositories;

namespace ZipStation.Worker.Services;

/// Mirror of the API's UserLookupService (no project reference between repos).
/// Resolves the customer's id in the external system via the project's UserLookup
/// settings and stores it on the customer record. Never throws.
public class UserLookupService
{
    private static readonly HttpClient _httpClient = new() { Timeout = TimeSpan.FromSeconds(10) };
    private readonly CustomerRepository _customerRepository;
    private readonly ILogger<UserLookupService> _logger;

    public UserLookupService(CustomerRepository customerRepository, ILogger<UserLookupService> logger)
    {
        _customerRepository = customerRepository;
        _logger = logger;
    }

    public async Task LookupAndStoreAsync(Project project, string? customerEmail)
    {
        try
        {
            var settings = project.Settings?.UserLookup;
            if (settings is not { Enabled: true } || string.IsNullOrWhiteSpace(settings.Url)
                || string.IsNullOrWhiteSpace(settings.UserIdField) || string.IsNullOrWhiteSpace(customerEmail))
                return;

            var customer = await _customerRepository.GetByEmailAndProjectAsync(customerEmail, project.Id);
            if (customer == null || !string.IsNullOrEmpty(customer.ExternalUserId))
                return;

            var externalUserId = await CallLookupEndpointAsync(settings, customerEmail);
            if (string.IsNullOrEmpty(externalUserId))
                return;

            customer.ExternalUserId = externalUserId;
            await _customerRepository.UpdateAsync(customer);
            _logger.LogInformation("Resolved external user id for customer {CustomerId} in project {ProjectId}", customer.Id, project.Id);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "User lookup failed for project {ProjectId}", project.Id);
        }
    }

    private static async Task<string?> CallLookupEndpointAsync(UserLookupSettings settings, string email)
    {
        var url = BuildUrl(settings.Url, email);
        using var request = new HttpRequestMessage(HttpMethod.Get, url);
        if (!string.IsNullOrEmpty(settings.AuthHeaderName) && !string.IsNullOrEmpty(settings.AuthHeaderValue))
            request.Headers.TryAddWithoutValidation(settings.AuthHeaderName, EncryptionHelper.Decrypt(settings.AuthHeaderValue));

        using var response = await _httpClient.SendAsync(request);
        response.EnsureSuccessStatusCode();

        var json = await response.Content.ReadAsStringAsync();
        return ExtractField(json, settings.UserIdField);
    }

    public static string BuildUrl(string urlTemplate, string email)
    {
        var encoded = Uri.EscapeDataString(email);
        if (urlTemplate.Contains("{email}", StringComparison.OrdinalIgnoreCase))
            return urlTemplate.Replace("{email}", encoded, StringComparison.OrdinalIgnoreCase);
        var separator = urlTemplate.Contains('?') ? '&' : '?';
        return $"{urlTemplate}{separator}email={encoded}";
    }

    /// Walks a dotted path (e.g. "userInfo.id") into the JSON. Numeric segments index arrays.
    public static string? ExtractField(string json, string path)
    {
        using var doc = JsonDocument.Parse(json);
        var element = doc.RootElement;
        foreach (var segment in path.Split('.', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (element.ValueKind == JsonValueKind.Array && int.TryParse(segment, out var index))
            {
                if (index < 0 || index >= element.GetArrayLength()) return null;
                element = element[index];
            }
            else if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(segment, out element))
            {
                return null;
            }
        }

        return element.ValueKind switch
        {
            JsonValueKind.String => element.GetString(),
            JsonValueKind.Number or JsonValueKind.True or JsonValueKind.False => element.GetRawText(),
            _ => null
        };
    }
}
