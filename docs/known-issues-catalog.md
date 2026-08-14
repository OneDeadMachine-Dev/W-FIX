# Known issues catalog / Каталог известных проблем

## Русский

W-Fix содержит fallback-каталог `Catalog/known-issues.json` и пытается получить два release asset:

- `known-issues.json`;
- `known-issues.json.sig` — Base64 ECDSA P-256/SHA-256 подпись точных байтов JSON.

Онлайн-файл принимается только при корректной подписи, поддерживаемой схеме, непросроченном `expiresAt`, уникальных
ID и HTTPS-ссылках на `learn.microsoft.com` либо `support.microsoft.com`. При ошибке используется последний валидный
подписанный кэш, затем встроенный fallback. Каталог не содержит PowerShell: `recommendedActionIds` может ссылаться
только на встроенные `legacy:*` действия.

Подпись выпуска:

```powershell
dotnet run --project tools/W-Fix.CatalogSigner -- sign `
  src/W-Fix.Core/Catalog/known-issues.json `
  path/to/known-issues-private.pem `
  artifacts/known-issues.json.sig
```

Закрытый ключ нельзя добавлять в Git. Публичный ключ находится в `known-issues-public.pem` и закреплён в приложении.
При ротации ключа сначала выпускается версия W-Fix с новым публичным ключом.

## English

W-Fix accepts a remote known-issues catalog only when its detached ECDSA P-256/SHA-256 signature is valid, the schema
is supported, the document is not expired, IDs are unique, and every source is an approved Microsoft HTTPS host.
Failure falls back to the last valid signed cache and then to the embedded catalog.

Catalog content is declarative and cannot deliver scripts or commands. It can only select `legacy:*` repair actions
already compiled into W-Fix. Never commit the private signing key; rotate the embedded public key through an application
release before signing the feed with a new key.
