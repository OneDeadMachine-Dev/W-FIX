using System.ComponentModel;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using WFix.Core.Abstractions;

namespace WFix.Core.Infrastructure;

public sealed class WindowsNetworkCredentialProvisioner : INetworkCredentialProvisioner
{
    private const int CredentialTypeDomainPassword = 2;
    private const int CredentialPersistLocalMachine = 2;
    private const int ErrorNotFound = 1168;

    public Task SaveForHostAsync(string hostName, NetworkCredential credential, CancellationToken cancellationToken = default)
    {
        var normalized = NormalizeHost(hostName);
        ArgumentNullException.ThrowIfNull(credential);
        if (string.IsNullOrWhiteSpace(credential.UserName) || string.IsNullOrEmpty(credential.Password))
            throw new ArgumentException("Для SMB-подключения нужны имя пользователя и непустой пароль.", nameof(credential));
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => Write(normalized, credential), cancellationToken);
    }

    public Task DeleteForHostAsync(string hostName, CancellationToken cancellationToken = default)
    {
        var normalized = NormalizeHost(hostName);
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() =>
        {
            if (CredDelete(normalized, CredentialTypeDomainPassword, 0)) return;
            var error = Marshal.GetLastWin32Error();
            if (error != ErrorNotFound) throw new Win32Exception(error);
        }, cancellationToken);
    }

    private static void Write(string hostName, NetworkCredential credential)
    {
        var bytes = Encoding.Unicode.GetBytes(credential.Password);
        if (bytes.Length > 512) throw new ArgumentException("Пароль превышает ограничение Windows Credential Manager.", nameof(credential));
        var pointer = Marshal.AllocCoTaskMem(bytes.Length);
        try
        {
            Marshal.Copy(bytes, 0, pointer, bytes.Length);
            var native = new NativeCredential
            {
                Type = CredentialTypeDomainPassword,
                TargetName = hostName,
                UserName = credential.UserName,
                CredentialBlob = pointer,
                CredentialBlobSize = bytes.Length,
                Persist = CredentialPersistLocalMachine,
                Comment = "W-Fix Pair Repair: authenticated SMB access to the selected printer host"
            };
            if (!CredWrite(ref native, 0)) throw new Win32Exception(Marshal.GetLastWin32Error());
        }
        finally
        {
            CryptographicOperations.ZeroMemory(bytes);
            Marshal.FreeCoTaskMem(pointer);
        }
    }

    private static string NormalizeHost(string hostName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(hostName);
        var value = hostName.Trim().TrimStart('\\').TrimEnd('.').ToUpperInvariant();
        if (value.Length is 0 or > 255 || value.Contains('\\') || value.Contains('/') || IPAddress.TryParse(value, out _))
            throw new ArgumentException("SMB credential должен быть привязан к имени компьютера, а не к IP или UNC-пути.", nameof(hostName));
        return value;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct NativeCredential
    {
        public int Flags;
        public int Type;
        [MarshalAs(UnmanagedType.LPWStr)] public string TargetName;
        [MarshalAs(UnmanagedType.LPWStr)] public string? Comment;
        public long LastWritten;
        public int CredentialBlobSize;
        public IntPtr CredentialBlob;
        public int Persist;
        public int AttributeCount;
        public IntPtr Attributes;
        [MarshalAs(UnmanagedType.LPWStr)] public string? TargetAlias;
        [MarshalAs(UnmanagedType.LPWStr)] public string UserName;
    }

    [DllImport("Advapi32.dll", EntryPoint = "CredWriteW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CredWrite([In] ref NativeCredential credential, int flags);

    [DllImport("Advapi32.dll", EntryPoint = "CredDeleteW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CredDelete(string target, int type, int flags);
}
