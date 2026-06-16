# Frage: Namensgleichheit Steuerliste / Parametertabelle (vf_ vs tf_)

## Kontext
SCR05 meldet: `Verfahren 'vf_ast_stea' nicht in Parametertabelle gefunden`.
Ursache: In der Steuerliste steht der Oracle-View-Name (`vf_...`), die
Parametertabelle ist aber auf den SQL-Zieltabellennamen (`tf_...`) geschluesselt.
Da Steuerliste und Parametertabelle beide vom Entwickler manuell gepflegt
werden, reicht es, in beiden den gleichen Namen zu verwenden.

---

## 🇩🇪 Nachricht

Hallo Feriz,

kurze Verständnisfrage: Die Steuerliste und die Parametertabelle werden ja
beide vom Entwickler gepflegt.

Können wir einfach den **gleichen Namen** wie in der Oracle-Quelle in **beiden**
(Steuerliste und Parametertabelle) verwenden? Dann passt die Prüfung
automatisch.

Dann müssten wir **im Code nichts ändern**, richtig?

Grüße

---

## 🇬🇧 Message

Quick question: the STL and the parameter table are both maintained by the
developer. Can we just use the **same name as the Oracle source** in **both**
the STL and the parameter table? Then the check matches automatically — and
**no code change** would be needed, right?
