SELECT id, status, is_default, path, start_time, last_event_time
FROM sys.traces;

EXEC sys.xp_readerrorlog 0, 1, NULL, NULL, '2026-10-06 01:30', '2026-10-06 04:00', 'asc';

SELECT x.value('(@timestamp)[1]', 'datetime2') AS zeit_utc,
       x.value('(data[@name="error_number"]/value)[1]', 'int') AS fehler,
       x.value('(data[@name="message"]/value)[1]', 'nvarchar(max)') AS meldung
FROM (SELECT CAST(event_data AS xml) AS d
      FROM sys.fn_xe_file_target_read_file('system_health*.xel', NULL, NULL, NULL)
      WHERE object_name = 'error_reported') e
CROSS APPLY e.d.nodes('/event') n(x)
WHERE x.value('(@timestamp)[1]', 'datetime2') BETWEEN '2026-10-05 23:00' AND '2026-10-06 03:00'
ORDER BY zeit_utc;

SELECT DB_NAME(database_id) AS db, name, physical_name, size / 128 AS size_mb
FROM sys.master_files
WHERE physical_name LIKE 'E:%'
ORDER BY size DESC;

SELECT bs.database_name, bs.backup_start_date, bs.backup_finish_date,
       bs.backup_size / 1048576 AS mb, bmf.physical_device_name
FROM msdb.dbo.backupset bs
JOIN msdb.dbo.backupmediafamily bmf ON bmf.media_set_id = bs.media_set_id
WHERE bs.backup_start_date BETWEEN '2026-10-05 22:00' AND '2026-10-06 06:00'
ORDER BY bs.backup_start_date;

SELECT SERVERPROPERTY('ErrorLogFileName') AS errorlog_pfad;
EXEC sys.xp_readerrorlog 0, 1, N'Trace';
EXEC sys.xp_readerrorlog 1, 1, N'Trace';
