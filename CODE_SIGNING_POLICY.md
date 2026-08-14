# Code signing policy / Политика подписи кода

Free code signing provided by [SignPath.io](https://signpath.io/), certificate by [SignPath Foundation](https://signpath.org/) — после одобрения проекта SignPath Foundation.

## Roles / Роли

- Committer and reviewer: [OneDeadMachine-Dev](https://github.com/OneDeadMachine-Dev)
- Signing approver: [OneDeadMachine-Dev](https://github.com/OneDeadMachine-Dev)

## Release policy / Политика выпуска

- Публичные binaries собираются только GitHub Actions из защищённой ветки `main` и опубликованного тега `v*`.
- Release workflow повторно выполняет restore, строгую Release-сборку, тесты и проверку зависимостей.
- Неподписанный artifact загружается в GitHub Actions до передачи в SignPath; SignPath проверяет происхождение сборки.
- Release signing требует ручного подтверждения. После подписи workflow проверяет Authenticode и timestamp, затем рассчитывает SHA-256.
- ZIP всегда создаётся из уже подписанного EXE.
- Detached ECDSA signature каталога известных проблем не заменяет Authenticode и управляется отдельным ключом.
- Закрытые ключи и API tokens находятся только в защищённых secret stores и не записываются в GitHub Release или логи.

- Public binaries are built only by GitHub Actions from protected `main` and a published `v*` tag.
- The release workflow repeats restore, strict Release build, tests, and dependency checks.
- The unsigned artifact is uploaded to GitHub Actions before submission so SignPath can verify build origin.
- Release signing requires manual approval. The workflow then verifies Authenticode and timestamp before calculating SHA-256.
- ZIP archives are always created from the signed EXE.
- The known-issues catalog ECDSA signature is independent from Authenticode and uses a separate key.
- Private keys and API tokens exist only in protected secret stores and are never written to releases or logs.

Until SignPath approval, public beta releases remain explicitly marked unsigned and include SHA-256 checksums. For internal testing, a separately documented self-signed certificate may be trusted manually on managed test computers; it is never presented as publicly trusted.

See also [Privacy Policy](PRIVACY.md) and [Security Policy](SECURITY.md).
