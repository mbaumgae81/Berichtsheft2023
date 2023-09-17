package dev.baumgaertner;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.temporal.IsoFields;
import java.time.temporal.WeekFields;
import java.util.List;
import java.util.Locale;
import java.util.Scanner;


import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import com.groupdocs.conversion.Converter;
import com.groupdocs.conversion.options.convert.PdfConvertOptions;

import javax.swing.*;


public class Main extends Thread {

    public static XWPFDocument berichtsheft;
    static JProgressBar status;
    public static XWPFDocument leseBheft;
    static Scanner read = new Scanner(System.in);
    private static File quelle = new File("./quelle/dokument.docx");
    String username = "Marco Baumgärtner";
    static int year = 2023;
    static int anzahl = 20;
    static int maxProgress = anzahl;
    static String neuerName = "";
    static Frame meinFrame = new JFrame("Berichtsheft 2023 ");
    static JPanel panel = new JPanel();
    static int progress = 0;
    static String name = "";
    static DateTimeFormatter format = DateTimeFormatter.ofPattern("dd.MM.yyyy");   // deutsches Datumsformat

    public static void main(String[] args) throws IOException, InvalidFormatException {
        /**
         *
         * Initialisierung mit Startwerten
         */
//
//

        Main thread = new Main();

        Frame meinFrame = new JFrame("Berichtsheft 2023 ");
        JPanel panel = new JPanel();
        meinFrame.setSize(400, 200);


        JTextField tfName = new JTextField("", 30);
        JButton buttonOK = new JButton("OK");
        JButton buttonExit = new JButton("Exit");

        panel.add(tfName);
        panel.add(new JLabel(checkForFolder() ? "Quell Ordner gefunden" : " Quell Ordner fehlt!"));
        panel.add(new JLabel("Bitte Vor und Nach namen eingeben"));
        panel.add(buttonExit);

        panel.add(buttonOK);
        panel.setLayout(new FlowLayout());
        status = new JProgressBar(0, maxProgress);
        panel.add(status);
        meinFrame.add(panel);
        System.out.println(maxProgress);

        meinFrame.setVisible(true);
        buttonOK.addActionListener(new ActionListener() {
                                       @Override
                                       public void actionPerformed(ActionEvent e) {
                                           neuerName = tfName.getText();
                                           name = neuerName;
                                           neuerName = neuerName.replace(" ", "_");
                                          thread.run();
                                       }
                                   }
        );
        buttonExit.addActionListener(new ActionListener() {
                                         @Override
                                         public void actionPerformed(ActionEvent e) {
                                             System.exit(1);
                                         }
                                     }
        );
    }

@Override
    public void run() {
        LocalDate date = LocalDate.now();
        /**
         * Anpassungen wenn von anderer Wochennummer erstellt wird
         */
        int weeknumber = 16;                                    //   Die Wochennummer bei der gestartet wird
        int nummerBerichtsheft = 43;                           // Bericht heft nummer oben rechts und erste stelle im dateinamen
        String trenn = "_";


        /*
        erstellen der dateien

         */

        for (int i = 0; i < anzahl; i++) {      //!!  einfach mal auf 22 wochen begrentzt !!


//            LocalDate endDate = date.with(IsoFields.WEEK_OF_WEEK_BASED_YEAR, weeknumber);       // Anhand der Wochennummer das ende der woche berechnen
            LocalDate startDate = getStartDate(year, weeknumber);
            LocalDate endDate = getEndtDate(year, weeknumber);

//            LocalDate startDate = endDate.minusDays(4);                             // Anhand des enddatums den anfang der woche berechnen
            String filename = nummerBerichtsheft + trenn + startDate + trenn + endDate + trenn + neuerName; // Dateinamen erstellen
            String readname = startDate + trenn + endDate + trenn + "MarcoBaumgärtner";             // erstelle name des berichtsheft von dem gelesen wird

//                System.out.println(filename);

            File ziel = new File("./temp/tmp.docx");
            try {// erstelle
                MyCopy.copyFile(quelle, ziel);                                                          // Temporäre datei
                berichtsheft = new XWPFDocument(OPCPackage.open("./temp/tmp.docx"));                       // lade Temporäre Datei
                leseBheft = new XWPFDocument(OPCPackage.open("./quelle/" + readname + ".docx"));      // lade Quell datei

            } catch (IOException a) {
                System.out.println("DAtei konnte nicht Kopiert werden rechte prüfen");
            } catch (InvalidFormatException e) {
                System.out.println("Datei konnte nicht geöffnet werden ");
            }

            List<XWPFTable> readTabelle = leseBheft.getTables();                                      // Lade Teabellen der quell Datei
            List<XWPFTable> tabelle = berichtsheft.getTables();                                       // Lade Tabellen der Zieldatei ( TMP )


            tabelle.get(0).getRow(0).getCell(1).setText(name);      // Tabelle 0 Zeile 0 Zelle 1 Eingegebener name wird eingetragen
            tabelle.get(0).getRow(0).getCell(4).setText(String.valueOf(nummerBerichtsheft));      // Tabelle 0 Zeile 0 Zelle 1 nummer des Berichtsheft wird eingetragen
            tabelle.get(0).getRow(2).getCell(1).setText(startDate.format(format));                  // Wochen Start wird eingetragen
            tabelle.get(0).getRow(2).getCell(3).setText(endDate.format(format));                    // Wochen ende wird eingetragen
            tabelle.get(1).getRow(0).getCell(0).setText(endDate.format(format));                    // Datum der Unterschrift wird eingetragen


            /**
             * Lese aus Quell Datei und schreibe in Ziel datei
             * Tabelle 0 get(0)
             * aktuelle Zeile getRow(x)
             * aktuelle Spalte ( Zelle ) getCell(y)
             *
             */
            tabelle.get(0).getRow(2).getCell(4).setText(readTabelle.get(0).getRow(2).getCell(4).getText()); // Trainer lesen und schreiben
            //
            tabelle.get(0).getRow(4).getCell(1).setText(readTabelle.get(0).getRow(4).getCell(1).getText()); // topics day 1
            tabelle.get(0).getRow(5).getCell(1).setText(readTabelle.get(0).getRow(5).getCell(1).getText()); // topics day 2
            tabelle.get(0).getRow(6).getCell(1).setText(readTabelle.get(0).getRow(6).getCell(1).getText()); // topics day 3
            tabelle.get(0).getRow(7).getCell(1).setText(readTabelle.get(0).getRow(7).getCell(1).getText()); // topics day 4
            tabelle.get(0).getRow(8).getCell(1).setText(readTabelle.get(0).getRow(8).getCell(1).getText()); // topics day 5

            tabelle.get(0).getRow(4).getCell(2).setText(readTabelle.get(0).getRow(4).getCell(2).getText()); // Lernfeld
            tabelle.get(0).getRow(5).getCell(2).setText(readTabelle.get(0).getRow(5).getCell(2).getText()); // Lernfeld
            tabelle.get(0).getRow(6).getCell(2).setText(readTabelle.get(0).getRow(6).getCell(2).getText()); // Lernfeld
            tabelle.get(0).getRow(7).getCell(2).setText(readTabelle.get(0).getRow(7).getCell(2).getText()); // Lernfeld
            tabelle.get(0).getRow(8).getCell(2).setText(readTabelle.get(0).getRow(8).getCell(2).getText()); // Lernfeld

            /**
             * Schreibe DokX Dateien
             * schließe alle Dateien
             */
            try (FileOutputStream out = new FileOutputStream("./temp/" + filename + ".docx")) {
                berichtsheft.write(out);
                berichtsheft.close();
                leseBheft.close();
                System.out.println(" Write " + i);
            } catch (IOException e) {
                System.out.println("Fehler beim schreiben der Datei");

            }

            /**
             *
             * Umwandlung der DocX in PDF
             *
             */
            String pdfname = "./pdf/" + filename + ".pdf";
            Converter converter = new Converter("./temp/" + filename + ".docx");
            converter.convert(pdfname, new PdfConvertOptions());

            // und weiter nächste woche :-)
            weeknumber++;
            nummerBerichtsheft++;
progress();
            status.setValue(progress);
            panel.updateUI();
            panel.revalidate();
            meinFrame.revalidate();
            meinFrame.repaint();

        }


    }

   static  public void progress(){
        progress ++;
        System.out.println(" aktueller Progress " + progress+ " max Progress :" + maxProgress);
        status.setValue(progress);
       System.out.println(status.getValue());
    }

    private static boolean checkForFolder() {
        Path tmp = Path.of("temp");
        Path pdf = Path.of("pdf");
        Path dokumentQuelle = Path.of("./quelle/dokument.docx");
        try {
            Files.createDirectories(pdf);
            Files.createDirectories(tmp);
        } catch (IOException e) {
            System.out.println(" Verzeichnis Existiert bereits");
        }

        if (Files.exists(dokumentQuelle)) {
            return true;
        }
        return false;
    }


    private static LocalDate getStartDate(int year, int wochennummer) {

        LocalDate startDate = LocalDate.of(year, 2, 1)
                .with(DayOfWeek.MONDAY)
                .with(WeekFields.of(Locale.GERMANY).weekOfWeekBasedYear(), wochennummer);

        return startDate;
    }

    private static LocalDate getEndtDate(int year, int wochennummer) {
        LocalDate endDate = LocalDate.of(year, 2, 1)
                .with(DayOfWeek.FRIDAY)
                .with(WeekFields.of(Locale.GERMANY).weekOfWeekBasedYear(), wochennummer);


        return endDate;
    }
}


