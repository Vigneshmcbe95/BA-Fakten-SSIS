-- Test SCR11-Datenprüfung: count(*) gegen rownum = 1 (Oracle, SQL Developer)
-- Schema, Tabelle, Partitionsspalte und Wert anpassen; jede Abfrage 2x ausführen, Zeiten notieren.

select count(*) from MSI_DM_ZWG.TF_AGZU_VORGANG partition for (20260900);
select 1 from MSI_DM_ZWG.TF_AGZU_VORGANG partition for (20260900) where rownum = 1;
select 1 from MSI_DM_ZWG.TF_AGZU_VORGANG where MON_ID = 20260900 and rownum = 1;
