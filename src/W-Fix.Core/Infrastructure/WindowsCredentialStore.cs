using System.ComponentModel;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using WFix.Core.Abstractions;
using WFix.Core.Models;

namespace WFix.Core.Infrastructure;

/// <summary>
/// Хранит альтернативные доменные учётные записи в Windows Credential Manager.
/// Секрет материализуется только на время создания удалённой сессии.
/// </summary>
public sealed class WindowsCredentialStore : ICredentialStore
{
    private const int CredentialTypeGeneric = 1;
    private const int CredentialPersistLocalMachine = 2;
    private const int ErrorNotFound = 1168;

    public Task SaveAsync(
        CredentialReference reference,
        NetworkCredential credential,
        CancellationToken cancellationToken = default)
    {
        Validate(reference);
        ArgumentNullException.ThrowIfNull(credential);
        cancellationToken.ThrowIfCancellationRequested();

        return Task.Run(() =>
        {
            var bytes = Encoding.Unicode.GetBytes(credential.Password ?? string.Empty);
            if (bytes.Length > 512)
                throw new ArgumentException("Пароль превышает ограничение Windows Credential Manager.", nameof(credential));

            var blob = Marshal.AllocCoTaskMem(bytes.Length);
            try
            {
                Marshal.Copy(bytes, 0, blob, bytes.Length);
                var native = new NativeCredential
                {
                    Type = CredentialTypeGeneric,
                    TargetName = reference.TargetName,
                    UserName = credential.UserName,
                    CredentialBlob = blob,
                    CredentialBlobSize = bytes.Length,
                    Persist = CredentialPersistLocalMachine
                };
                if (!CredWrite(ref native, 0))
                    throw new Win32Exception(Marshal.GetLastWin32Error());
            }
            finally
            {
                if (bytes.Length > 0)
                    Array.Clear(bytes);
                Marshal.FreeCoTaskMem(blob);
            }
        }, cancellationToken);
    }

    public Task<NetworkCredential?> ReadAsync(
        CredentialReference reference,
        CancellationToken cancellationToken = default)
    {
        Validate(reference);
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run<NetworkCredential?>(() =>
        {
            if (!CredRead(reference.TargetName, CredentialTypeGeneric, 0, out var pointer))
            {
                var error = Marshal.GetLastWin32Error();
                if (error == ErrorNotFound)
                    return null;
                throw new Win32Exception(error);
            }

            try
            {
                var native = Marshal.PtrToStructure<NativeCredential>(pointer);
                var password = native.CredentialBlobSize == 0
                    ? string.Empty
                    : Marshal.PtrToStringUni(native.CredentialBlob, native.CredentialBlobSize / 2) ?? string.Empty;
                return new NetworkCredential(native.UserName, password);
            }
            finally
            {
                CredFree(pointer);
            }
        }, cancellationToken);
    }

    public Task DeleteAsync(
        CredentialReference reference,
        CancellationToken cancellationToken = default)
    {
        Validate(reference);
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() =>
        {
            if (CredDelete(reference.TargetName, CredentialTypeGeneric, 0))
                return;
            var error = Marshal.GetLastWin32Error();
            if (error != ErrorNotFound)
                throw new Win32Exception(error);
        }, cancellationToken);
    }

    private static void Validate(CredentialReference reference)
    {
        ArgumentNullException.ThrowIfNull(reference);
        if (!reference.TargetName.StartsWith("W-Fix/", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Credential target должен начинаться с 'W-Fix/'.", nameof(reference));
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

    [DllImport("Advapi32.dll", EntryPoint = "CredReadW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CredRead(string target, int type, int reservedFlag, out IntPtr credentialPtr);

    [DllImport("Advapi32.dll", EntryPoint = "CredDeleteW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CredDelete(string target, int type, int flags);

    [DllImport("Advapi32.dll", SetLastError = true)]
    private static extern void CredFree(IntPtr credential);
}
