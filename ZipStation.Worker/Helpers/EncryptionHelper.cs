using System.Security.Cryptography;
using System.Text;

namespace ZipStation.Worker.Helpers;

public static class EncryptionHelper
{
    private static string? _key;

    public static void Initialize(string? encryptionKey)
    {
        if (!string.IsNullOrEmpty(encryptionKey))
        {
            _key = Convert.ToBase64String(SHA256.HashData(Encoding.UTF8.GetBytes(encryptionKey)));
        }
    }

    public static bool IsInitialized => !string.IsNullOrEmpty(_key);

    public static string Decrypt(string cipherText)
    {
        if (string.IsNullOrEmpty(cipherText)) return cipherText;
        if (!cipherText.StartsWith("ENC:")) return cipherText;
        if (!IsInitialized) return cipherText;

        var fullBytes = Convert.FromBase64String(cipherText[4..]);

        using var aes = Aes.Create();
        aes.Key = Convert.FromBase64String(_key!);

        var iv = new byte[16];
        Buffer.BlockCopy(fullBytes, 0, iv, 0, 16);
        aes.IV = iv;

        var cipherBytes = new byte[fullBytes.Length - 16];
        Buffer.BlockCopy(fullBytes, 16, cipherBytes, 0, cipherBytes.Length);

        using var decryptor = aes.CreateDecryptor();
        var plainBytes = decryptor.TransformFinalBlock(cipherBytes, 0, cipherBytes.Length);

        return Encoding.UTF8.GetString(plainBytes);
    }
}
