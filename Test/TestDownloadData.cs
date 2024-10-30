using BergNotenWASM.Model;

namespace BergNotenWASM.Test
{
    public class TestDownloadData
    {
        public static readonly List<Teilnehmer> TestTeilnehmer =
        [
            new() { Vorname = "Leonardo", Nachname = "DiCaprio", Geburtsdatum = new DateTime(1974, 11, 11), Verein = "Hollywood Club" },
            new() { Vorname = "Scarlett", Nachname = "Johansson", Geburtsdatum = new DateTime(1984, 11, 22), Verein = "Avengers Club" },
            new() { Vorname = "Tom", Nachname = "Hanks", Geburtsdatum = new DateTime(1956, 7, 9), Verein = "Forest Gump Fans" },
            new() { Vorname = "Emma", Nachname = "Watson", Geburtsdatum = new DateTime(1990, 4, 15), Verein = "Hogwarts Alumni" },
            new() { Vorname = "Brad", Nachname = "Pitt", Geburtsdatum = new DateTime(1963, 12, 18), Verein = "Ocean's Eleven" },
            new() { Vorname = "Angelina", Nachname = "Jolie", Geburtsdatum = new DateTime(1975, 6, 4), Verein = "Humanitarian League" },
            new() { Vorname = "Will", Nachname = "Smith", Geburtsdatum = new DateTime(1968, 9, 25), Verein = "Men in Black Association" },
            new() { Vorname = "Jennifer", Nachname = "Lawrence", Geburtsdatum = new DateTime(1990, 8, 15), Verein = "Hunger Games Squad" },
            new() { Vorname = "Robert", Nachname = "Downey Jr.", Geburtsdatum = new DateTime(1965, 4, 4), Verein = "Iron Man Legends" },
            new() { Vorname = "Chris", Nachname = "Hemsworth", Geburtsdatum = new DateTime(1983, 8, 11), Verein = "Thor's Warriors" }
        ];
        
        public static readonly List<Pruefungen> TestPruefungen =
        [
            new()
            {
                Name = "Grundschwung",
                Beschreibung = "Prüfung des Grundschwungs auf blauen Pisten, Fokus auf Technik und Stabilität."
            },
            new()
            {
                Name = "Schrägfahrt",
                Beschreibung =
                    "Schrägfahrt-Prüfung für das Gleichgewicht und die Kanteneinstellung auf mittleren Hängen."
            },
            new()
            {
                Name = "Parallelschwung",
                Beschreibung = "Prüfung des Parallelschwungs auf roten Pisten mit Fokus auf Kontrolle und Rhythmus."
            },
            new()
            {
                Name = "Kanteneinsatz",
                Beschreibung =
                    "Prüfung des präzisen Kanteneinsatzes auf anspruchsvollen Pisten, Bewertung der Genauigkeit und Balance."
            },
            new()
            {
                Name = "Carving",
                Beschreibung = "Carving-Prüfung auf harten Pisten, Geschwindigkeit und Technik werden geprüft."
            },
            new()
            {
                Name = "Buckelpisten",
                Beschreibung = "Prüfung auf Buckelpisten, bewertet werden Rhythmus und Körperspannung."
            },
            new()
            {
                Name = "Kurvenfahren",
                Beschreibung =
                    "Kurventechnik auf schmalen, steilen Pisten, Schwerpunkt auf Kontrolle und Belastungswechsel."
            },
            new()
            {
                Name = "Off-Piste",
                Beschreibung =
                    "Off-Piste-Technik auf tief verschneiten Hängen, Prüfung auf Anpassung und Geländebeherrschung."
            },
            new()
            {
                Name = "Schussfahrt",
                Beschreibung =
                    "Sicherheitsprüfung der Schussfahrt auf einer geraden Piste, Augenmerk auf Körperhaltung und Bremsfähigkeit."
            },
            new()
            {
                Name = "Freestyle-Skiing",
                Beschreibung = "Freestyle-Prüfung mit Sprüngen und Tricks, Fokus auf Kreativität und Körperkontrolle."
            }
        ];

    }
}
