# Partitionstausch `tf_bb_bewa_zkt` – zusätzlicher Index, Rückfrage

**Betreff:** Partitionstausch `tf_bb_bewa_zkt` – zusätzlicher Index gefunden, kurze Rückfrage

Hallo zusammen,

beim Partitionstausch von **`tf_bb_bewa_zkt`** bricht SQL Server ab mit:
> *"There is no identical index in source table … for the index **`nci_bak_zsk_mow`**."*

**Was wir gefunden haben:**
- Auf der Faktentabelle gibt es **zwei** Indizes: den **CCI** (Clustered Columnstore – den legt der Ladeprozess an, so wie in der Parametertabelle `FaktenClusteredIndex = CCI` definiert) **und** einen zusätzlichen **UNIQUE Nonclustered Index `nci_bak_zsk_mow`** auf `(bak_id, zsk_id, mow_id)`.
- Dieser zweite Index ist **nicht** in der Parametertabelle hinterlegt und wird **nicht** vom Ladeprozess erzeugt. Die Parametertabelle sagt dem Ladeprozess nur, **welchen** Index er anlegen soll – sie kennt aber **nicht** Indizes, die bereits zusätzlich auf der Faktentabelle existieren.
- Für einen Partitionstausch müssen Quelle (Staging) und Ziel (Faktentabelle) **dieselben** Indizes haben. Da die Staging-Tabelle `nci_bak_zsk_mow` nicht hat, lehnt SQL Server den SWITCH ab.
- Gut: der Index ist **partition-aligned** – er kann also problemlos auf der Staging-Tabelle nachgebaut werden, sobald wir ihn kennen.

**Frage an euch:**
Wann bzw. wie wurde **`nci_bak_zsk_mow`** angelegt – per **separatem Skript** oder **manuell**? Stammt er evtl. noch aus dem **alten Ladeverfahren**? Wir müssen das wissen, damit wir den Index im Ladeprozess korrekt mit auf die Staging-Tabellen übernehmen können.

Danke!

---

**Subject:** Partition switch `tf_bb_bewa_zkt` – extra index found, quick question

Hi all,

The partition switch for **`tf_bb_bewa_zkt`** fails with:
> *"There is no identical index in source table … for the index **`nci_bak_zsk_mow`**."*

**What we found:**
- The fact table has **two** indexes: the **CCI** (created by the load, as set in the parameter table `FaktenClusteredIndex = CCI`) **and** an extra **UNIQUE nonclustered index `nci_bak_zsk_mow`** on `(bak_id, zsk_id, mow_id)`.
- That second index is **not** in the parameter table and is **not** created by the load. The parameter table only tells the load **which** index to create — it does **not** know about indexes that already exist on the fact table.
- A partition switch requires source (staging) and target (fact table) to have the **same** indexes. The staging table doesn't have `nci_bak_zsk_mow`, so SQL Server rejects the SWITCH.
- Good news: the index **is partition-aligned**, so we can simply recreate it on the staging tables once we know about it.

**Question:**
When/how was **`nci_bak_zsk_mow`** created — by a **separate script** or **manually**? Is it left over from the **old loading method**? We need to know so we can replicate it correctly on the staging tables in the load.

Thanks!
