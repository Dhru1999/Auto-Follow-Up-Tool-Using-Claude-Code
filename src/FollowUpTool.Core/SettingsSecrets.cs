namespace FollowUpTool.Core;

/// <summary>Encrypts/decrypts small secrets. The app uses Windows DPAPI; tests use a fake.</summary>
public interface ISecretProtector
{
    byte[] Protect(byte[] plain);
    byte[] Unprotect(byte[] cipher);
}

/// <summary>
/// Converts secrets to and from the Base64 strings stored in settings.json.
/// Never throws on unreadable input: a secret that cannot be decrypted (for
/// example settings copied from another Windows account) comes back empty so
/// the user is simply asked to enter it again.
/// </summary>
public static class SettingsSecrets
{
    public static string Encrypt(ISecretProtector protector, string? plainText)
    {
        if (string.IsNullOrEmpty(plainText)) return "";
        var cipher = protector.Protect(System.Text.Encoding.UTF8.GetBytes(plainText));
        return Convert.ToBase64String(cipher);
    }

    public static string Decrypt(ISecretProtector protector, string? encryptedBase64)
    {
        if (string.IsNullOrEmpty(encryptedBase64)) return "";
        try
        {
            var plain = protector.Unprotect(Convert.FromBase64String(encryptedBase64));
            return System.Text.Encoding.UTF8.GetString(plain);
        }
        catch (Exception ex) when (ex is FormatException or System.Security.Cryptography.CryptographicException)
        {
            return "";
        }
    }
}
