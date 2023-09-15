package dev.baumgaertner.probe;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.temporal.IsoFields;

public class main {
    public static void main(String[] args) {
        int weeknumber = 15;
        DateTimeFormatter format = DateTimeFormatter.ofPattern("dd.MM.yyyy");   // deutsches Datumsformat
        LocalDate date = LocalDate.now();
        LocalDate endDate = date.with(IsoFields.WEEK_OF_WEEK_BASED_YEAR, weeknumber);       // Anhand der Wochennummer das ende der woche berechnen
        LocalDate startDate = endDate.minusDays(4);

        System.out.println(weeknumber);
        System.out.println(date.format(format));
        System.out.println(endDate.format(format));

    }
}
