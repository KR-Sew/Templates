# Предоставление удалённого доступа сотруднику ГК "Везу.Ру"
![Win ACME Version](https://img.shields.io/badge/Win_ACME-v2.2.0-blue)  
![Let's Encrypt](https://img.shields.io/badge/Powered_by-Let's_Encrypt-brightgreen)

## Порядок действий
1. **Добавьте сотрудника в доменные группы**  
   - В группу доступа `PW.Salova\RDG Workstation Users` для доступа к уд.р.столу рабочего ПК
   - В группу доступа `PW.Salova\RDG Users` для доступа к шлюзу уд.р.столов  

2. **Добавьте сотрудника в локальную группу**  
   - В локальную группу `Пользователи удалённого рабочего стола` для доступа к уд.р.столу рабочего ПК

3. **Добавьте ПК сотрудника в доменную группу**  
   - В группу  `PW.Salova\RDG Workstations` для которых разрешены удалённый подключения из внешней сети.  
4. **Настройте брандмауэр Windows**  
   - проверьте правила разрешающие подключения к удалённому рабочему столу в настройках `Брандмауэр Windows`
5. **Включите/Проверьте "Удалённый рабочий стол"**  
   - Проверьте влючен `"Удалёный рабочий стол"`
   ```powershell
   Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server' -Name "fDenyTSConnections" | Select-Object -ExpandProperty fDenyTSConnections
   # Returns 0 = включен, 1 = выключен
   ```
   - Включить "Удалённый рабочий стол"
   ```powershell


