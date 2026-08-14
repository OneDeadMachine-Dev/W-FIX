# Политика конфиденциальности / Privacy Policy

## Русский

W-Fix не содержит телеметрии, рекламы и автоматической отправки отчётов. Приложение работает локально или с компьютерами, которые явно выбрал оператор.

- Логи хранятся локально в `%LocalAppData%\W-Fix\Logs` и могут содержать имя локального компьютера и пользователя.
- Отчёты о ремонте хранятся в `%ProgramData%\W-Fix\Runs` и могут содержать имена выбранных компьютеров, принтеров, драйверов и результаты действий.
- Обезличенный support bundle создаётся только по команде пользователя. Имена компьютеров и принтеров заменяются псевдонимами; пароли, ссылки Credential Manager, пути снимков и содержимое документов печати не экспортируются.
- Альтернативные учётные данные сохраняются только после явного согласия в Windows Credential Manager. Пароли не записываются в конфигурацию, аргументы процессов, логи и отчёты.
- При диагностике W-Fix может загрузить декларативный каталог известных проблем с официального GitHub Release проекта. Загруженный каталог проверяется цифровой подписью и не может содержать исполняемый код.
- W-Fix не передаёт сведения третьим лицам, если пользователь сам не экспортировал и не отправил отчёт или support bundle.
- Pair Repair обменивается данными только между двумя явно выбранными ПК. Приглашение и offline-снимок не содержат паролей; live-listener и временное правило Firewall удаляются после сессии.

Сообщить о проблеме конфиденциальности можно через [Private vulnerability reporting](https://github.com/OneDeadMachine-Dev/W-FIX/security/advisories/new).

## English

W-Fix contains no telemetry, advertising, or automatic report uploads. It operates locally or against computers explicitly selected by the operator.

- Logs are stored locally in `%LocalAppData%\W-Fix\Logs` and may contain the local computer and user names.
- Repair reports are stored in `%ProgramData%\W-Fix\Runs` and may contain selected computer, printer, and driver names plus action results.
- A sanitized support bundle is created only on explicit request. Computer and printer names are replaced with aliases; passwords, Credential Manager references, snapshot paths, and print-document contents are excluded.
- Alternate credentials are saved only with explicit consent in Windows Credential Manager. Passwords are never written to configuration, process arguments, logs, or reports.
- During diagnostics W-Fix may download a declarative known-issues catalog from the project's official GitHub Release. The catalog is signature-verified and cannot contain executable code.
- W-Fix sends no information to third parties unless the user explicitly exports and shares a report or support bundle.
- Pair Repair exchanges data only between two explicitly selected PCs. Invitations and offline snapshots contain no passwords; the live listener and temporary Firewall rule are removed after the session.

Report a privacy concern through [Private vulnerability reporting](https://github.com/OneDeadMachine-Dev/W-FIX/security/advisories/new).
