# Namensregel: Steuerliste (STL) vs. Parametertabelle

## 🇩🇪 E-Mail

**Betreff: Faktenladen – welcher Name wird wo verwendet (STL vs. Parametertabelle)**

Hallo zusammen,

kurz und einfach, damit das Laden sauber durchläuft:

**1. Steuerliste (STL):**
Hier steht der **Name des Objektes auf Oracle** (die Quelle, z. B. View
`vf_stea` oder Tabelle `tf_bzg_bezugsgroesse`). Mit diesem Namen liest das
Paket die Daten aus Oracle.

**2. Parametertabelle:**
Hier steht unter **`Faktentabelle`** der **Name der Zieltabelle auf SQL Server**
(z. B. `tf_stea`). Mit diesem Namen wird auf dem Zielserver alles gebaut:
Stagingtabellen (`_in_`, `_out_`), Indizes, Partitionsfunktion und der
Partition-Switch in die Faktentabelle. Ebenfalls kommt von hier die
**Partitionsspalte** (`Faktenpartitionsspalte`).

**3. Wo es zusammenpassen muss (wichtig):**
Der **Name aus der Steuerliste = der Verfahrensname** und genau dieser muss
als **Zeile in der Parametertabelle** vorhanden sein. Über diesen Namen findet
das Paket die passenden Parameter (Zieltabelle, Partitionsspalte usw.).

→ Kurz: **STL-Name = Schlüssel in der Parametertabelle.**
   Stimmt er nicht überein, kommt „Verfahren nicht in Parametertabelle gefunden“.

**Beispiel `stea`:**
- STL: `vf_stea`  (Oracle-Quelle)
- Parametertabelle: Zeile `Verfahren = vf_stea`, `Faktentabelle = tf_stea`,
  `Faktenpartitionsspalte = mon_id`
- Ablauf: lese Oracle `vf_stea` → lade nach `tf_stea_in_<Monat>` → Index →
  Switch in `tf_stea`.

**Beispiel Normalfall (`bzg`):** Oracle-Objekt und Zieltabelle heißen gleich –
in STL und Parametertabelle steht überall `tf_bzg_bezugsgroesse`.

Viele Grüße

---

## 🇬🇧 Summary

- **STL** = Oracle **source** object name (used to read from Oracle).
- **Parameter table `Faktentabelle`** = SQL Server **target** table name (used for
  staging `_in_`/`_out_`, indexes, partition function, switch into the fact table);
  the **partition column** also comes from here.
- **Match point:** the **STL name = Verfahren key**, and that exact name must
  exist as a **row in the parameter table**. If it doesn't match →
  "Verfahren nicht in Parametertabelle gefunden".

Example `stea`: STL `vf_stea` → parameter row `Verfahren=vf_stea`,
`Faktentabelle=tf_stea` → read Oracle `vf_stea`, load into `tf_stea`.
Normal case `bzg`: same name everywhere (`tf_bzg_bezugsgroesse`).
