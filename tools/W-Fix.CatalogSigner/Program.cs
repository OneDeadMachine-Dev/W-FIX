using System.Security.Cryptography;

if (args.Length == 0)
    return Usage();

try
{
    return args[0].ToLowerInvariant() switch
    {
        "generate" when args.Length == 3 => Generate(args[1], args[2]),
        "sign" when args.Length == 4 => Sign(args[1], args[2], args[3]),
        "verify" when args.Length == 4 => Verify(args[1], args[2], args[3]),
        _ => Usage()
    };
}
catch (Exception ex)
{
    Console.Error.WriteLine(ex.Message);
    return 1;
}

static int Generate(string privateKeyPath, string publicKeyPath)
{
    RefuseOverwrite(privateKeyPath);
    RefuseOverwrite(publicKeyPath);
    using var key = ECDsa.Create(ECCurve.NamedCurves.nistP256);
    File.WriteAllText(privateKeyPath, key.ExportPkcs8PrivateKeyPem());
    File.WriteAllText(publicKeyPath, key.ExportSubjectPublicKeyInfoPem());
    Console.WriteLine($"Ключи созданы. Закрытый ключ не добавляйте в Git: {Path.GetFullPath(privateKeyPath)}");
    return 0;
}

static int Sign(string catalogPath, string privateKeyPath, string signaturePath)
{
    var catalog = File.ReadAllBytes(catalogPath);
    using var key = ECDsa.Create();
    key.ImportFromPem(File.ReadAllText(privateKeyPath));
    var signature = key.SignData(catalog, HashAlgorithmName.SHA256);
    File.WriteAllText(signaturePath, Convert.ToBase64String(signature));
    Console.WriteLine($"Подпись создана: {Path.GetFullPath(signaturePath)}");
    return 0;
}

static int Verify(string catalogPath, string publicKeyPath, string signaturePath)
{
    var catalog = File.ReadAllBytes(catalogPath);
    var signature = Convert.FromBase64String(File.ReadAllText(signaturePath).Trim());
    using var key = ECDsa.Create();
    key.ImportFromPem(File.ReadAllText(publicKeyPath));
    var valid = key.VerifyData(catalog, signature, HashAlgorithmName.SHA256);
    Console.WriteLine(valid ? "VALID" : "INVALID");
    return valid ? 0 : 2;
}

static void RefuseOverwrite(string path)
{
    if (File.Exists(path))
        throw new InvalidOperationException($"Файл уже существует: {Path.GetFullPath(path)}");
}

static int Usage()
{
    Console.Error.WriteLine("W-Fix.CatalogSigner generate <private.pem> <public.pem>");
    Console.Error.WriteLine("W-Fix.CatalogSigner sign <catalog.json> <private.pem> <catalog.sig>");
    Console.Error.WriteLine("W-Fix.CatalogSigner verify <catalog.json> <public.pem> <catalog.sig>");
    return 64;
}
