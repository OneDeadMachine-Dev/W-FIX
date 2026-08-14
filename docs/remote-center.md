# Remote Repair Center / Центр удалённого ремонта

## Русский

### Требования

- доменные Windows 10/11 x64;
- административная учётная запись;
- WinRM/PowerShell Remoting;
- Task Scheduler и вошедший пользователь — только для HKCU/принтера по умолчанию;
- запуск W-Fix от имени администратора.

### Рабочий процесс

1. Откройте **Remote Center** из главного окна.
2. Добавьте FQDN/NetBIOS вручную или найдите машины в Active Directory.
3. По умолчанию используется текущая Windows-учётка и Kerberos. Альтернативную доменную учётку можно явно
   сохранить в Credential Manager; пароль сразу очищается из формы.
4. Нажмите **Preflight**. Недоступная capability отключает только зависящие от неё действия.
5. Нажмите **Диагностика** и изучите факты, уверенность и план.
6. Нажмите **Ремонт**. Обычный план требует одного подтверждения; необратимые шаги — дополнительного.
7. Откройте HTML/JSON-отчёт либо экспортируйте обезличенный ZIP.

Параллелизм ограничен диапазоном 1–10, значение по умолчанию — 3. Ошибка одной машины не останавливает пакет.
Отмена прекращает запуск новых целей и инициирует rollback незавершённой обратимой цели. Перезагрузка никогда не
отправляется автоматически: в конце используется отдельная кнопка и подтверждение.

## English

Remote Center officially targets domain-joined Windows 10/11 x64 clients. Add computers manually or through Active
Directory, run preflight, capture inventory, review evidence and the proposed repair plan, then confirm the batch.

The default authentication path is the current Windows identity with Kerberos. An alternate domain account is stored
only after an explicit action in Windows Credential Manager. Concurrency defaults to three and can be set from one to
ten. A failed target is verified and rolled back independently; it does not stop the remaining targets. Reboots always
require a separate end-of-run confirmation.
