/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.todivefor.processgarmin;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.SwingWorker;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.write.DateTime;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/**
 *
 * @author peterream
 */
public class ProcessDataRows extends SwingWorker 
{
    public int position;
    public String inputString;
    
    public ProcessDataRows()
    {

    }

    @Override
    protected Object doInBackground() throws Exception
    {
        final String COURSENAME = "CourseName";              // CourseName tag
        final String COURSEDATE = "CourseDate";              // CourseDate tag
        final String HOLESCORE = "HoleScore";                // HoleScore tag
        
        String[] filesInDir;
        WritableWorkbook workbook;                           // Workbook to process
        String sheetName = "First Sheet";                    // Sheet name
        WritableCellFormat dateCellFormatMDY;
        WritableSheet sheet = null;                                // Sheet to process
  
        File aDirectory = new File(ProcessGarmin.inputPath);        // create a file that is really a directory              

        filesInDir = aDirectory.list();                             // get a listing of all files in the directory
        
        Arrays.sort(filesInDir);                                    // sort the list of files (optional)
 
//      Create Spreadsheet
        
        workbook = null;
        try
        {
            workbook = Workbook.createWorkbook(new File(ProcessGarmin.outputPath + "/" +
                    ProcessGarmin.spreadsheetName));                                  // Set spreadsheet Name
            sheet = workbook.createSheet(sheetName, 0);
        }
        catch (IOException ex)
        {
            Logger.getLogger(ProcessDataRows.class.getName()).log(Level.SEVERE, null, ex);
        }
        
//      Setup cell formats for date
 
        dateCellFormatMDY = new WritableCellFormat
            (new jxl.write.DateFormat("mm/dd/yy"));                     // Used to format date
        WritableCellFormat dateCellFormatA = new WritableCellFormat();  // Used for cell alignment right
        
        
//      Write header rows to spreadsheet (Course, date, and 18holes)
        
        int col = 0;                                    // Column 0
        int row = 0;                                    // row 0 header data
        Label COURSELabel = new Label(col, 
            row, "Course");                             // Course header (0,0)
        col = 1;                                        // Next column
        try                                             // Right align date header
        {
            dateCellFormatA.setAlignment(Alignment.RIGHT);  // Format right "Date"
        }
        catch (WriteException ex)
        {
            Logger.getLogger(ProcessDataRows.class.getName()).log(Level.SEVERE, null, ex);
        }
        Label DATELabel = new Label(col,
            row, "Date", dateCellFormatA);              // Date header (0,1)
        try                                             // Write course and date header cells
            {
                sheet.addCell(COURSELabel);             // Add course header
                sheet.addCell(DATELabel);               // add date header
            }
            catch (WriteException ex)
            {
                Logger.getLogger(ProcessDataRows.class.getName()).log(Level.SEVERE, null, ex);
            }
        
//      Add holes 1 - 18 header

        int hole = 0;                                   // Start hole #
        for (hole = 1; hole < 19; hole++)               // Loop thru 18 holes
            {
                col = hole+1;                           // Hole score column
                jxl.write.Number number = new jxl.write.Number(col, 
                        row, hole);                     // To spreadsheet (0,n)
                try                                     // Write hole number header cells
                {
                    sheet.addCell(number);              // Add hole number
                }
                catch (WriteException ex)
                {
                    Logger.getLogger(ProcessDataRows.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
        
//      Process files within directory and extract data

        row = 2;                                                // First data row
        boolean extract = false;                                    // Extract data or not
        int i;
        for (i = 0; i < filesInDir.length; i++ )      // Loop through files
        {
            if (ProcessGarmin.yearToExtract.equals("") || 
                    filesInDir[i].contains
                    (ProcessGarmin.yearToExtract))                  // File contain year or no year entered?
                extract = true;                                     // Yes, set to extract
            if (ProcessGarmin.courseToExtract.equals("") || 
                    filesInDir[i].contains
                    (ProcessGarmin.courseToExtract))                // Course to extract or no course entered?
                extract = extract && true;                          // Yes, set to extract anded with year result
            else                                                    // No
                extract = false;                                    // Set to not extract
            if (ProcessGarmin.debug)
            {
                System.out.println( "file: " + filesInDir[i] );
            }
            ProcessGarmin.scoreCard = ProcessGarmin.inputPath + "/" + filesInDir[i];    // Path + filename
            if (extract)                                                // Year and course valid?
            {
                if (ProcessGarmin.debug)
                    
                    System.out.println("Processing " + (i + 1) + " of " +
                            filesInDir.length + " - " + filesInDir[i]);         // Yes processing
                try                                                     // Yes
                {
                    buildInputString();                                 // Read file into inputString
                }
                catch (IOException ex)
                {
                    Logger.getLogger(ProcessDataRows.class.getName()).log(Level.SEVERE, null, ex);
                }
                
//              Get course name from input string

                position = 0;                                               // Position within inputString
                String courseName = getTaggedInfo(position, COURSENAME);    // Get course name

//              Get date from input string

                String courseDate = getTaggedInfo(position, COURSEDATE);    // Get course date
                DateFormat df = new SimpleDateFormat("MMM dd, yyyy");       // Format of date
                Date scoreDate = null;                                      // Score date object
                try
                {
                    scoreDate = df.parse(courseDate);                       // Convert date to date object
                }
                catch (ParseException ex)
                {
                    Logger.getLogger(ProcessDataRows.class.getName()).log(Level.SEVERE, null, ex);
                }

//              Get hole scores from input string

                int iHole;                                      // Hole index
                int[] holeScore = new int[18];                  // Hole array
                boolean skipRound = false;                      // Skip round or not
                for (iHole = 0; iHole < 18; iHole++)            // Loop through 18 hole scores
                {
                    String sHoleScore = getTaggedInfo(position, HOLESCORE); // Get score by hole
                    holeScore[iHole] = Integer.parseInt(sHoleScore); // Convert score to integer
                    if (holeScore[iHole] == 255)                // Score for hole?
                    {
                        skipRound = true;                       // Yes, skip this round
                    }
                }
                if (ProcessGarmin.debug)
                {
                    System.out.println("Course: "+ courseName +
                            "\nDate: " + courseDate);
                    for (iHole = 0; iHole < 18; iHole++)
                    {
                        System.out.println("Hole " + (iHole + 1) +
                                ": " + holeScore[iHole]);
                    }
                }

//              Write data rows to spreadsheet

                col = 0;                                    // Course column
                Label courseNameLabel = new Label(col,
                        row, courseName);                   // Assume good name
                col = 1;                                    // Date column
                DateTime date = new DateTime(col,
                        row, scoreDate, dateCellFormatMDY); // Score date to excel
                if (skipRound)                              // Skip round?
                {
                    col = 0;                                // Yes, course column
                    courseNameLabel = new Label(col, row,
                            courseName + " MISSING SCORE"); // Mark as skipped
                }
                try
                {
                    sheet.addCell(courseNameLabel);         // Add course name
                    sheet.addCell(date);                    // Add date
                }
                catch (WriteException ex)
                {
                    Logger.getLogger(ProcessDataRows.class.getName()).
                            log(Level.SEVERE, null, ex);
                }
                //                    int hole;                 // Hole index
                for (hole = 0; hole < 18; hole++)           // Loop thru 18 holes
                {
                    col = hole+2;                               // Hole score column
                    jxl.write.Number number = new jxl.write.Number(col,
                            row, holeScore[hole]);              // To spreadsheet
                    try
                    {
                        sheet.addCell(number);                  // Add score
                    }
                    catch (WriteException ex)
                    {
                        Logger.getLogger(ProcessDataRows.class.getName()).log(Level.SEVERE, null, ex);
                    }
                }
                row++;                                      // Increment sheet row
            }                                               // End extract
            else                                            // Data not to be extracted
            {
                if (ProcessGarmin.debug)
                    System.out.println("Bypassing  " + (i + 1) + " of " +
                            filesInDir.length + " - " +
                            filesInDir[i]);                     // Yes processing
            }
            
//          Pass progress back to ProcessGarmin (Event Dispatch Thread)

            int percentDone;
            double numFiles = filesInDir.length;
            double curFile = i;
            if (ProcessGarmin.debug)
                System.out.println(i);                              // File counter
            percentDone = (int) Math.floor((curFile / numFiles) * 100);
            
//          Set progres in 10% increments
            
            if (percentDone % 5 == 0)
                setProgress(percentDone);                           // Progress
        }                                                           // End for
        
//      All done, close workbook
        
        try
        {
            workbook.write();                                       // Write to spreadsheet
            workbook.close();                                       // Close spreadsheet
        }
        catch (IOException | WriteException ex)
        {
            Logger.getLogger(ProcessDataRows.class.getName()).
                    log(Level.SEVERE, null, ex);
        }
        if (ProcessGarmin.debug)
            System.out.println("doInBackground() finished");
        return (row - 2);                                           // Number of records processed
    }
    
    public void buildInputString() throws IOException
    {
        inputString = null;                                             // Reset inputString
        int count = 0;                                                  // Reset byte count
        try
            (FileInputStream fileIn = new FileInputStream(ProcessGarmin.scoreCard))       // Input Stream
        {
            boolean eof = false;                                    // New file turn off eof
            int readIt;                                     // Byte read
            while (!eof)
            {
                readIt = fileIn.read();                     // Read byte
                if (readIt == -1)                           // EOF?
                {
                    eof = true;                             // Yes, set eof
                    break;                                  // Exit if
                }
                char c = (char) readIt;                     // Convert to character             
                inputString = inputString + c;              // Build inputString
                count++;                                    // Characters read
            }
            fileIn.close();                                 // Close input file
        }
        catch (Exception e)
        {
            System.out.println(e);
        }
        if (ProcessGarmin.debug)
        {
            System.out.println(inputString);
        }
        if (ProcessGarmin.debug)
            System.out.println(count + " bytes processed");                                    // Close input
    }
        private String getTaggedInfo(int startPosition, String findString)
    {
        String foundTagged = null;                                          // Tagged item
        String startTag = "<" + findString + ">";                           // Start tag
        int startIndex = inputString.indexOf(startTag, startPosition);      // Index of start tag
        startIndex = startIndex + startTag.length();                        // Beginning of tagged data
        String endTag = "</" + findString + ">";                            // End tag
        int endIndex = inputString.indexOf(endTag, startIndex);             // Index of end tag
        foundTagged = inputString.substring(startIndex, endIndex);          // Tagged data
        position = endIndex + endTag.length();                              // Point past end tag
        return foundTagged;                                                 // Return found data
        
    }
}
