using Amazon.S3;
using Amazon.S3.Model;
using ZipStation.Worker.Entities;
using ZipStation.Worker.Helpers;

namespace ZipStation.Worker.Services;

public class FileStorageService
{
    private readonly ILogger<FileStorageService> _logger;

    public FileStorageService(ILogger<FileStorageService> logger)
    {
        _logger = logger;
    }

    public async Task<string> UploadAsync(FileStorageSettings settings, string storageKey, Stream stream, string contentType)
    {
        using var client = CreateClient(settings);
        var request = new PutObjectRequest
        {
            BucketName = settings.BucketName,
            Key = storageKey,
            InputStream = stream,
            ContentType = contentType
        };
        await client.PutObjectAsync(request);
        _logger.LogInformation("Uploaded file {StorageKey} to bucket {Bucket}", storageKey, settings.BucketName);
        return storageKey;
    }

    private static AmazonS3Client CreateClient(FileStorageSettings settings)
    {
        var keyId = EncryptionHelper.Decrypt(settings.KeyId);
        var appKey = EncryptionHelper.Decrypt(settings.AppKey);
        var config = new AmazonS3Config
        {
            ServiceURL = settings.Endpoint,
            ForcePathStyle = true
        };
        return new AmazonS3Client(keyId, appKey, config);
    }
}
