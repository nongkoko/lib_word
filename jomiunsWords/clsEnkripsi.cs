using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

public class Encryption64
{
    public static string Encrypt(string stringToEncrypt, string sEncryptionKey)
    {
        var aReturn = string.Empty;
        var key = Encoding.UTF8.GetBytes(sEncryptionKey.Length > 8 ? sEncryptionKey.Substring(0, 8) : sEncryptionKey.PadRight(8, '\0'));
        var IV = new byte[] { 0x12, 0x34, 0x56, 0x78, 0x90, 0xAB, 0xCD, 0xEF };
        using (var des = new DESCryptoServiceProvider())
        {
            var inputByteArray = Encoding.UTF8.GetBytes(stringToEncrypt);
            using (var ms = new MemoryStream())
            {
                using (var cs = new CryptoStream(ms, des.CreateEncryptor(key, IV), CryptoStreamMode.Write))
                {
                    cs.Write(inputByteArray, 0, inputByteArray.Length);
                    cs.FlushFinalBlock();
                    aReturn = Convert.ToBase64String(ms.ToArray());
                }
            }
        }
        return aReturn;
    }

    public static string Decrypt(string stringToDecrypt, string sEncryptionKey)
    {
        var IV = new byte[] { 0x12, 0x34, 0x56, 0x78, 0x90, 0xAB, 0xCD, 0xEF };

        if (string.IsNullOrEmpty(stringToDecrypt))
            return "";

        var key = Encoding.UTF8.GetBytes(sEncryptionKey.Length > 8 ? sEncryptionKey.Substring(0, 8) : sEncryptionKey.PadRight(8, '\0'));
        using (var des = new DESCryptoServiceProvider())
        {
            var inputByteArray = Convert.FromBase64String(stringToDecrypt);
            using (var ms = new MemoryStream())
            {
                using (var cs = new CryptoStream(ms, des.CreateDecryptor(key, IV), CryptoStreamMode.Write))
                {
                    cs.Write(inputByteArray, 0, inputByteArray.Length);
                    cs.FlushFinalBlock();
                    return Encoding.UTF8.GetString(ms.ToArray());
                }
            }
        }
    }
}