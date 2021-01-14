# Einleitung 

Der CovidTestOutlookCollector hilft Ihnen dabei, die Testergebnisse, die
Eurofins LifeCodexx Ihnen schickt, automatisch auszuwerten.

Die Software ist lizensiert unter den GNU GPL v3.0 und der Quellcode ist
einsehbar unter <https://github.com/Wulfheart/CovidTestSammler>.

# Funktionsweise

Eurofins LifeCodexx sendet Ihnen pro Test mindestens zwei Emails zu.
Eine enthält eine passwortgeschützte PDF-Datei im Anhang und eine andere
enthält das Passwort für die PDF-Datei. Berichten zufolge ist es auch
schon passiert, dass auch das Emailpaar mehrfach gesendet wurde. Wenn
Sie eine große Anzahl an Testergebnissen haben, weil Sie beispielsweise
ein Einrichtungsleiter in einer Einrichtung sind, in der Reihentestungen
(die Sie natürlich mit dem Reihentestungstool
<https://github.com/Wulfheart/CovidBulkBavaria> erstellt haben )
anstanden, dann kann es schnell mühsam werden, alle Testergebnisse
auszuwerten und einen Überblick zu bekommen. Sie müssen für jede Person
einzeln die Email mit dem Passwort finden, die Email mit dem Ergebnis
finden, den Anhang öffnen, das Passwort einfügen, das Testergebnis
speichern, sicherstellen, dass Sie kein Testergebnis doppelt haben und
dann noch eine Übersicht erstellen. Mit dem CovidTestOutlookCollector
schaffen Sie das in kürzester Zeit. Testweise wurden über zweihundert
Testergebnisse in 1,5 Minuten zugeordnet.

# Installation

Der Installationsvorgang setzt grundlegende technische Kenntnisse wie
Herunterladen von ZIP-Dateien und Entpacken derselbigen voraus.

## Installation der .NET 5.0 Runtime

Bitte stellen Sie sicher, dass Sie die .NET Runtime 5.0 installiert
haben. Falls sie nicht installiert ist, laden Sie es sich unter
<https://dotnet.microsoft.com/download/dotnet/5.0> herunter. Achten Sie
darauf, die .NET Runtime 5.0.0 auszuwählen. Führen Sie daraufhin die
heruntergeladene Datei aus.

## Installation des CovidTestOutlookCollectors

Laden Sie unter
[https://github.com/Wulfheart/CovidTestSammler/releases/latest](https://github.com/Wulfheart/CovidBulkBavaria/releases/latest)
die Datei mit Namen „CovidTestOutlookCollectorSetup.exe“ herunter und
führen Sie sie lokal aus. Danach sollte eine Anwendung mit der
Bezeichnung „CovidTestOutlookCollector“ auf Desktop und Startmenü
verfügbar sein.

Falls bei der Installation Schwierigkeiten auftreten sollten, nehmen Sie
bitte entweder über die Github Issues oder die Emailadresse Kontakt auf.

Falls bei Programmstart eine Meldung auftritt, dass eine .NET Runtime
nicht installiert ist, folgen Sie bitte nicht dem angegebenen Link,
sondern dem unter Punkt „Installation der .NET 5.0 Runtime“ in diesem
Dokument.

# Weitere Voraussetzungen

Sie benötigen Microsoft Outlook als Emailprogramm. Getestet wurde es mit
Outlook 2019 und 365, aber das Programm sollte auch mit älteren
Outlookversionen funktionieren.

# Empfohlener Ablauf

## Exportieren der Emails aus Outlook

Zuerst exportieren Sie die Emails, die Sie von der Teststelle kriegen,
in einen Ordner. Dieser Ordner sollte immer derselbe sein. So wird
sichergestellt, dass kein Testergebnis übersehen wird. Sie exportieren
die Emails aus Outlook, indem Sie die Emails markieren (ähnlich Dateien
im Explorer) und dann in den gewünschten Ordner ziehen. Falls eine Email
bereits in dem Zielordner sein sollte, dann wird Windows eine Meldung
anzeigen und Sie können diese Email überspringen, sodass sie nicht in
das Ziel kopiert wird. Falls eine Email aus irgendwelchen Gründen
doppelt in dem Ordner sein sollte: Keine Sorge, der
CovidTestOutlookCollector erkennt das und kümmert sich darum. Jede
Test-ID (vom Barcode) wird nur einmal ausgewertet.

Tipp: Normalerweise kommen die Emails mit den Testergebnissen immer von
der gleichen Emailadresse. Sie können mit einer Weiterleitungsregel in
Outlook die Emails automatisch in einen Ordner in Outlook weiterleiten
lassen. Damit haben Sie dann alle Emails, die mit Covidtestergebnissen
zu tun haben, direkt an einer Stelle. Falls Sie nicht wissen, wie Sie
Weiterleitungsregeln in Outlook erstellen, lesen Sie bitte Anleitungen
im Internet oder fragen Sie Ihren Techniker oder Systemadministrator.

## Auswertung

Starten Sie den CovidTestOutlookCollector und klicken Sie auf den Button
„Auswertung starten“. Wählen Sie in dem sich öffnenden Fenster den
Ordner, in dem die Mails (.msg-Dateien) sind. Klicken Sie dann auf
„Ordner auswählen“. Die Auswertung startet daraufhin automatisch.

Dann beginnt das Programm zu arbeiten. Falls Fehler auftreten, werden
Sie ausgegeben. Zusätzlich wird unten im Log auch gezeigt, was genau das
Programm gerade macht. Falls Sie sich mit Fehlern an den Entwickler
wenden, inkludieren Sie bitte diese Meldungen.

Wenn das Programm fertig ist, dann erscheint ein Dialogfenster und
Lognachricht mit dem Text „Vorgang abgeschlossen. Die Daten können in
dem Ordner /Pfad/zu/ihren/Emails/PDFs/JJJJ\_MM\_TT\_SS\_mm\_ss abgerufen
werden“. In diesem Ordner finden sich alle in diesem Durchlauf neu
ausgewerteten Testergebnisse einzeln als PDF, ein Dokument mit all
diesen PDFs zusammengefasst sowie eine Exceldatei zur Übersicht.
