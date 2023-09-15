package dev.baumgaertner;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.temporal.IsoFields;
import java.util.List;
import java.util.Scanner;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import com.groupdocs.conversion.Converter;
import com.groupdocs.conversion.options.convert.PdfConvertOptions;



    public class Main {
        public static XWPFDocument berichtsheft ;
        public static XWPFDocument leseBheft;
        static Scanner read = new Scanner(System.in);
        public static void main(String[] args) throws IOException, InvalidFormatException {
            /**
             *
             * Initialisierung mit Startwerten
             */
            File quelle = new File("quelle/dokument.docx");
            String username = "Marco Baumgärtner";
            System.out.println("Bitte vor und nachnamen angeben :");
            String neuerName = read.nextLine();

            DateTimeFormatter format = DateTimeFormatter.ofPattern("dd.MM.yyyy");
            LocalDate date = LocalDate.now();
            int weeknumber = 16;
            int nummerBheft = 43;
            String trenn = "_";


        /*
        erstellen der dateien
         */
            for (int i = 0; i < 22; i++) {


                LocalDate endDate = date.with(IsoFields.WEEK_OF_WEEK_BASED_YEAR, weeknumber);
                LocalDate startDate = endDate.minusDays(4);
                String filename = nummerBheft + trenn + startDate + trenn + endDate + trenn + neuerName;

                String readname = startDate + trenn + endDate + trenn + "MarcoBaumgärtner";
                System.out.println(filename);

                File ziel = new File("tmp.docx");
                MyCopy.copyFile(quelle, ziel);

                berichtsheft = new XWPFDocument(OPCPackage.open("tmp.docx"));
                leseBheft = new XWPFDocument(OPCPackage.open("quelle/" + readname + ".docx"));

                List<XWPFTable> readTabelle = leseBheft.getTables();
                List<XWPFTable> tabelle = berichtsheft.getTables();




                tabelle.get(0).getRow(0).getCell(1).setText(neuerName);      // Tabelle 0 Zeile 0 Zelle 1
                tabelle.get(0).getRow(0).getCell(4).setText(String.valueOf(nummerBheft));      // Tabelle 0 Zeile 0 Zelle 1
                tabelle.get(0).getRow(2).getCell(1).setText(startDate.format(format));
                tabelle.get(0).getRow(2).getCell(3).setText(endDate.format(format));
                tabelle.get(1).getRow(0).getCell(0).setText(endDate.format(format));


                tabelle.get(0).getRow(2).getCell(4).setText(readTabelle.get(0).getRow(2).getCell(4).getText()); // Trainer
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

                try (FileOutputStream out = new FileOutputStream(filename + ".docx")) {
                    berichtsheft.write(out);
                    berichtsheft.close();
                    leseBheft.close();
                }

                String pdfname = filename+".pdf";

                Converter converter = new Converter(filename + ".docx");
                converter.convert(pdfname, new PdfConvertOptions());



                weeknumber++;
                nummerBheft++;
            }


        }


    }