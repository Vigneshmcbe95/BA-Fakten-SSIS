-- Prueft, ob der Master Key automatisch ueber den Service Master Key geoeffnet wird (1 = ja).
USE msi_dm_zwg;
SELECT name, is_master_key_encrypted_by_server
FROM sys.databases WHERE name = DB_NAME();
