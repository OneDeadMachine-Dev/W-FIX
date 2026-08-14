# Безопасность / Security

## Поддерживаемые версии / Supported versions

Security fixes are provided for the latest stable release and the current public beta. Older releases should be upgraded before reporting a vulnerability.

## Сообщение об уязвимости / Reporting a vulnerability

Не публикуйте уязвимости, секреты, доменные имена и журналы организации в обычных GitHub Issues. Используйте [Private vulnerability reporting](https://github.com/OneDeadMachine-Dev/W-FIX/security/advisories/new).

Do not disclose vulnerabilities, secrets, domain names, or organizational logs in public GitHub Issues. Use [Private vulnerability reporting](https://github.com/OneDeadMachine-Dev/W-FIX/security/advisories/new).

В отчёте укажите версию W-Fix, версию Windows, ожидаемое и фактическое поведение и минимальные шаги воспроизведения. Прикладывайте только обезличенный support bundle. Мы подтвердим получение через GitHub Security Advisory и не будем просить пароль или содержимое документов печати.

Include the W-Fix version, Windows version, expected and actual behavior, and minimal reproduction steps. Attach only a sanitized support bundle. We will acknowledge the report through GitHub Security Advisories and will never request a password or print-document contents.

## Границы доверия / Trust boundaries

- W-Fix требует административных прав, потому что изменяет службы, реестр, политики печати и драйверы.
- Remote Center выполняет действия только на явно выбранных целях.
- Pair Repair использует TLS 1.2+, временный ECDSA-ключ, pinning, одинаковый код и подтверждение на обоих ПК. По каналу разрешены только типизированные DTO и встроенные `pair.*` действия.
- Pair listener не открывается на Public-профиле; SMB1 не поддерживается.
- Загружаемый known-issues catalog декларативен, подписан ECDSA и ссылается только на встроенные действия.
- Private keys, GitHub tokens and SignPath tokens must never be committed to the repository.
