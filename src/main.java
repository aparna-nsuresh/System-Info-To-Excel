import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Scanner;
import jxl.*;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WritableCellFormat;
import jxl.write.WriteException;

/**
 * VPS Excel Management Utility
 * SUPPLEMENT TO THE COMPUTER INFORMATION BATCH FILE
 * Version 2.05 - Uses imported JXL Library for Excel development
 * Updated by Sanjeet Bagga on 7/8/2016
 * Note - Excel Documents will be 97-2003 Excel (.xls) files.
 */

public class main {

    public static void main(String[] args) {
        addToExcel();

    }

    /**
     * Adds data from text files from a specified folder into a specified excel document
     * Uses a reference Excel Document and produces an output Excel Document with data organized
     */

    public static void addToExcel()
    {
        String current = System.getProperty("user.dir");
        File folder = new File(current + "\\output");
        File[] listFiles = folder.listFiles();

        Workbook workbook1 = null;

        //Creating the format excel sheet that can be used for quick format editing on the document
        try {
            workbook1 = Workbook.getWorkbook(new File(current + "\\output\\format.xls"));
        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        } finally {
        }

        /**
         * Creating the output excel sheet that all the data will be put into
         */
        WritableWorkbook copy = null;
        try {
            copy = Workbook.createWorkbook(new File(current + "\\output\\output.xls"), workbook1);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
        }
        WritableSheet sheet = copy.getSheet(0);

        /**
         *Searches through all the text files and gathers information to be put into the output excel file
         */
        for (File file : listFiles) {
            try {
                int column = 0;
                int row = sheet.getRows();
                row++;
                Scanner s = new Scanner(file);
                while (s.hasNextLine()) {
                    String str = s.nextLine();


                    column++;
                    WritableCellFormat cellFormat = new WritableCellFormat();
                    try {
                        cellFormat.setWrap(true);
                    } catch (WriteException e) {
                        e.printStackTrace();
                    }

                    /**
                     * Removes all useless information from the output file
                     */
                    if (str.contains("    IP Address:                           "))
                    {
                            str = str.substring(42);

                    }
                    else if(str.contains("OS Name:                   ") || str.contains("Original Install Date:     "))
                    {
                        str = str.substring(27);
                    }
                    else if(str.contains("Total # of free bytes        : "))
                    {
                        str = str.substring(31);
                    }
                    else if(str.contains("Total # of bytes             : ") || str.contains("Total # of avail free bytes  : "))
                    {
                        str = str.substring(31);

                    }

                    Label label = new Label(column, row, str, cellFormat);

                    String test = str;
                    int r = row;

                    /**
                     *Wraps all the programs into a single cell
                     */

                    while(s.hasNext() && (str.contains("Name") && s.next().contains("Name")))
                    {
                        test += "\n" + s.nextLine();


                    }

                    /**
                     * Finalizes changes and puts the information in the cells
                     */
                    System.out.print(test);
                    label.setString(test);
                        try {
                            sheet.addCell(label);
                        } catch (WriteException e) {
                            e.printStackTrace();
                        }
                    }

                s.close();
            }
            catch (FileNotFoundException e) {
                e.printStackTrace();

            }
        }
        /**
         * Saves the changes into the excel document
         */
        try {
            copy.write();
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            copy.close();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        }
    }
}