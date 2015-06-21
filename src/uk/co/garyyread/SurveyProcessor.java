package uk.co.garyyread;

import java.io.File;
import java.io.IOException;
import java.time.DateTimeException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.HashMap;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/**
 * Survey processor, bespoke application developed for a masters university 
 * student at Swansea university in aid to processes years worth of data.
 * @version 3.1
 * @author Gary Read
 * @since 2015
 */
public class SurveyProcessor {
    private boolean debug;

    //Non-static vars
    private Workbook workBook;
    private WritableWorkbook resultBook;
    private final String TAB;
    private int ROW_START;
    private int BEACH; 
    private int ID; 
    private int DATE; 
    private int JULIAN_DATE; 
    private int AGE_CLASS; 
    private Sheet workingSheet;

    /**
     * Public class constructor...
     */
    public SurveyProcessor() {
        this.debug = false;
        this.TAB = "\t";
        this.ROW_START = 1;
        this.BEACH = 0;
        this.ID = 1;
        this.DATE = 2;
        this.JULIAN_DATE = 3;
        this.AGE_CLASS = 4;
    }

    /**
     * Read and assign a workbook from a string file name
     *
     * @param loc string file name.
     * @return True = file set, false is a fail.
     */
    public boolean loadWorkbook(String loc) {
        boolean result = false;

        try {
            workBook = Workbook.getWorkbook(new File(loc));

            if (debug) {
                debug("readWorkBook(" + loc + ")");
            }
            result = true;

        } catch (IOException ex) {
            displayMessage("File not recognised.");
            Logger.getLogger(SurveyProcessor.class.getName()).log(Level.SEVERE, null, ex);
        } catch (BiffException ex) {
            displayMessage("Workbook not read.");
            Logger.getLogger(SurveyProcessor.class.getName()).log(Level.SEVERE, null, ex);
        }

        return result;
    }
    
        public boolean loadWorkbook(File file) {
        boolean result = false;

        try {
            workBook = Workbook.getWorkbook(file);

            if (debug) {
                debug("readWorkBook(" + file.getAbsolutePath() + ")");
            }
            result = true;

        } catch (IOException ex) {
            displayMessage("File not recognised.");
            Logger.getLogger(SurveyProcessor.class.getName()).log(Level.SEVERE, null, ex);
        } catch (BiffException ex) {
            displayMessage("Workbook not read.");
            Logger.getLogger(SurveyProcessor.class.getName()).log(Level.SEVERE, null, ex);
        }

        return result;
    }

    /**
     * Create a workbook to write results too
     *
     * @param fileName File name of new writable workbook
     * @return reference to workbook
     * @throws IOException
     */
    public WritableWorkbook createWritableWorkbook(String fileName) throws IOException {
        WritableWorkbook res = Workbook.createWorkbook(new File(fileName));
        this.resultBook = res;
        return resultBook;
    }

    /**
     * Business logic...
     *
     * @param name Sheet name from workbook to work from.
     */
    public void processSheet(String name) throws WriteException, IOException {
        //Get sheet
        this.workingSheet = workBook.getSheet(name);

        //Check sheet exists
        if (workingSheet == null) {
            displayMessage("Sheet " + name + " doesn't exist.");
            if (debug) {
                debug("EXCEPTION:processSheet(" + name + "),sheet does not exist.");
            }
        }

        //Get columns from working sheet
        int ID_POS = 0;
        int BEACH_POS = 1;
        int AGE_CLASS_POS = 2;
        int DATE_POS = 3;
        int JULIAN_DATE_POS = 4;
        Cell[][] data = new Cell[5][];
        data[ID_POS] = workingSheet.getColumn(ID);
        data[BEACH_POS] = workingSheet.getColumn(BEACH);
        data[AGE_CLASS_POS] = workingSheet.getColumn(AGE_CLASS);
        data[DATE_POS] = workingSheet.getColumn(DATE);
        data[JULIAN_DATE_POS] = workingSheet.getColumn(JULIAN_DATE);

        //debugging info - print held data
        if (debug) {
            int i;
            for (i = 0; i < data[0].length; i++) {
                try {
                    debug(i + TAB + data[0][i].getContents() + TAB + data[1][i].getContents() + TAB + data[1][i].getCellFormat().getBackgroundColour().getValue() + TAB + data[2][i].getContents() + TAB + data[3][i].getContents() + TAB + data[4][i].getContents());
                } catch (NullPointerException ex) {
                    debug("EXCEPTION:processSheet(" + name + "),row:" + (i+1) + " is null");
                } catch (ArrayIndexOutOfBoundsException ex) {
                    debug("EXCEPTION:processSheet(" + name + "),array is out of bounds at " + i);
                }
            }
            debug("processSheet(" + name + ")" + ",row count:" + data[0].length);
        }

        //HashMap stored data as <KEY, String[] {BEACH,ID,COLOUR,C0D,C0J,C1D,C1J,C2D,C2J,C3D,C3J,C4D,C4J,C5D,C5J}>
        HashMap<String, String[]> map = new HashMap<>();
        int AGE_CLASS_OFFSET = 3; //array off-set for age class
        int JULIAN_AGE_CLASS_OFFSET = 9; //array off-set for age class for julian date
        for (int i = ROW_START; i < data[ID_POS].length; i++) {

            //Collect row data
            try {
                String idStr = data[ID_POS][i].getContents();
                String idColourStr = "" + data[ID_POS][i].getCellFormat().getBackgroundColour().getValue();
                String beachStr = data[BEACH_POS][i].getContents();
                int ageClass = sanatizeAgeClassInput(data[AGE_CLASS_POS][i].getContents());
                LocalDate date = convertStrToDate(data[DATE_POS][i].getContents(), i);
                String julianDate = data[JULIAN_DATE_POS][i].getContents();

                //Anything greater than 5 is bad!
                if (ageClass > 5) {
                    displayMessage("FATAL ERROR in Sheet \"" + workingSheet.getName() + "\" at row[" + (i+1) + "]"
                            + "\n" + "Please fix error in Class column!"
                            + "\n" + "Failed to convert text to date: \"" + ageClass + "\"");
                    debug("processSheet failed as ageClass was out of range - EXIT");
                    return;

                    //Compare date if key exists (find eariest date) then add to HashMap
                } else if (map.containsKey(idStr)) {
                    String[] arrOld = map.get(idStr);
                    String dateStrOld = arrOld[ageClass + AGE_CLASS_OFFSET];
                    LocalDate dateOld = null;
                    if (dateStrOld.length() > 0 && ageClass != -1) {
                        dateOld = LocalDate.parse(dateStrOld);
                    }

                    //update colour and beach
                    String[] arrNew = arrOld.clone();
                    arrNew[0] = beachStr;
                    arrNew[2] = idColourStr;
                    if (dateOld == null && ageClass > -1) {
                        arrNew[ageClass + AGE_CLASS_OFFSET] = date.toString();
                        arrNew[ageClass + JULIAN_AGE_CLASS_OFFSET] = julianDate;

                        //update map
                        map.put(idStr, arrNew);

                        if (debug) {
                            debug("ADDED TO EXISTING" + TAB + Arrays.toString(arrOld) + " to " + Arrays.toString(arrNew));
                        }

                        //Compare dates and add the earliest occurance of the seal
                    } else if (ageClass > -1) {
                        //Replace date if we find an ealier one
                        if (dateOld.compareTo(date) > 0) {
                            arrNew[ageClass + AGE_CLASS_OFFSET] = date.toString();
                            arrNew[ageClass + JULIAN_AGE_CLASS_OFFSET] = julianDate;

                            //update map
                            map.put(idStr, arrNew);

                            if (debug) {
                                debug("UPDATED" + TAB + Arrays.toString(arrOld) + " to " + Arrays.toString(arrNew));
                            }
                        }
                    }

                    //New item, add straight to the array
                } else {
                    String[] arr = {beachStr, idStr, idColourStr, "", "", "", "", "", "", "", "", "", "", "", ""};
                    if (ageClass != -1) {
                        arr[ageClass + AGE_CLASS_OFFSET] = date.toString();
                        arr[ageClass + JULIAN_AGE_CLASS_OFFSET] = julianDate;
                    }
                    map.put(idStr, arr);

                    if (debug) {
                        debug("NEW" + TAB + Arrays.toString(arr));
                    }
                }
            } catch (NullPointerException ex) {
                if (debug) {
                    debug("EXCEPTION:processSheet(" + name + "),missing data from row... row skipped... Null pointer exception");
                }
            } catch (DateTimeException ex) {
                if (debug) {
                    debug("EXCEPTION:processSheet(" + name + "),missing data from row... row skipped... Date Format exception");
                }
                return;
            } catch (ArrayIndexOutOfBoundsException ex) {
                if (debug) {
                    debug("EXCEPTION:processSheet(" + name + "),missing data from row... row skipped... Array out of bounds exception");
                }
            }
        }

        //debugging
        int i = 1;
        for (i = 1; i <= map.size() + 1 && debug; i++) {
            String[] arr = map.get(i + "");
            debug("id:" + i + " " + Arrays.toString(arr));
        }
        
        writeResultsToWorkbook(name, map);
    }

    /**
     * Write data from processSheet to a WritableWorkbook
     *
     * @param name Name to call sheet in the working workbook
     * @param map Data to write to the sheet
     */
    private void writeResultsToWorkbook(String name, HashMap<String, String[]> map) throws WriteException, IOException {
        //Create work sheet
        createWritableWorkbook(name + "_result.xls");
        WritableSheet ws = resultBook.createSheet(name, 0);
        
        //Write column headers
        ws.addCell(new Label(0, 0, "Beach"));
        ws.addCell(new Label(1, 0, "Pup ID"));
        ws.addCell(new Label(2, 0, "C0"));
        ws.addCell(new Label(3, 0, "C1"));
        ws.addCell(new Label(4, 0, "C2"));
        ws.addCell(new Label(5, 0, "C3"));
        ws.addCell(new Label(6, 0, "C4"));
        ws.addCell(new Label(7, 0, "C5"));
        ws.addCell(new Label(8, 0, "C0"));
        ws.addCell(new Label(9, 0, "C1"));
        ws.addCell(new Label(10, 0, "C2"));
        ws.addCell(new Label(11, 0, "C3"));
        ws.addCell(new Label(12, 0, "C4"));
        ws.addCell(new Label(13, 0, "C5"));

        //Loop tnough rows of data
        for (int r = 1; r <= map.size() + 1; r++) {
            String[] arr = map.get(r + "");

            //Add each column of data from the row 'r'
            for (int c = 0; arr != null && c < arr.length; c++) {
                int altC = c - 1;
                String cont = arr[c];

                if (c == 1) {
                    //Get cell colour and create cell format
                    int colourValue = Integer.parseInt(arr[c + 1]);
                    WritableCellFormat format = new WritableCellFormat();
                    format.setBorder(Border.ALL, BorderLineStyle.THIN, Colour.GRAY_25);
                    if (colourValue == 192) {
                        format.setBackground(Colour.UNKNOWN);
                    } else {
                        format.setBackground(Colour.getInternalColour(colourValue));
                    }

                    try {
                        jxl.write.Number nc = new jxl.write.Number(c, r, Double.parseDouble(cont), format);
                        ws.addCell(nc);
                    } catch (NumberFormatException ex) {
                        if (debug) {
                            debug("EXCEPTION:method.1(),No age class found in data \"" + cont + "\"");
                        }
                    }

                    //date cell
                } else if (c > 2 && c < 9) {
                    try {
                        if (cont.length() > 0) {
                            ws.addCell(new Label(altC, r, cont));
                        }
                    } catch (StringIndexOutOfBoundsException ex) {
                        if (debug) {
                            debug("EXCEPTION:method.2(),No date found in data \"" + cont + "\"");
                        }
                    }

                    //Julian cell
                } else if (c > 8) {
                    try {
                        jxl.write.Number nc = new jxl.write.Number(altC, r, Double.parseDouble(cont));
                        ws.addCell(nc);
                    } catch (NumberFormatException ex) {
                        if (debug) {
                            debug("EXCEPTION:method.3(),No julian date found in data \"" + cont + "\"");
                        }
                    }

                    //All other columns (minus colour information)
                } else {
                    if (c != 2) {
                        ws.addCell(new Label(c, r, cont));
                    }
                }
            }
        }

        //Write data and close workbook
        resultBook.write();
        resultBook.close();
    }

    /**
     * Not all cell values are regular, take first integer from string (always
     * number)
     *
     * @param ageClassStr String value indicating.
     * @return -1 for invalid input, 0 to 5 otherwise.
     */
    private int sanatizeAgeClassInput(String ageClassStr) {
        for (int i = 0; i < ageClassStr.length(); i++) {
            try {
                //return Integer.parseInt("" + ageClassStr.charAt(ageClassStr.length() - 1));
                return Integer.parseInt(ageClassStr.substring(i,i+1));
            } catch (Exception e) {
                if (debug) debug("sanatizeAgeClassInput(" + ageClassStr + ") failed at least once.");
            }
        }
        
        return -1;
    }

    /**
     * Fix excel converting error (though JEXCEL).
     *
     * @param dateStr Date entered as a string.
     * @return LocalDate object representing the correct date from the survey.
     */
    private LocalDate convertStrToDate(String dateStr, int row) {
        //convert date to working date type
        if (debug) {
            debug("convertStrToDate(" + dateStr + ")");
        }
        DateTimeFormatter dateFormat = DateTimeFormatter.ofPattern("dd/MM/yy");
        LocalDate date = null;

        try {
            date = LocalDate.parse(dateStr, dateFormat);
        } catch (DateTimeException ex) {
            displayMessage("FATAL ERROR in Sheet \"" + workingSheet.getName() + "\" at row[" + (row+1) + "]"
                    + "\n" + "Please fix error in date column!"
                    + "\n" + "Failed to convert text to date: \"" + dateStr + "\"");
            debug("convertStrToDate Failed - EXIT");
            throw ex;
        }

        int year_error_range = 2025;
        if (date.getYear() > year_error_range) {
            int yearError = 100;
            int yearCorrected = date.getYear() - yearError;
            date = date.withYear(yearCorrected);
        }
        return date;
    }

    /**
     * Method to flip debugging mode
     *
     * @param debugging
     */
    public void setDebugging(boolean debugging) {
        debug = debugging;
    }

    /**
     * Return list of sheet names that can be worked on.
     *
     * @return String array of workbook sheets.
     */
    public String[] getSheetNames() {
        return workBook.getSheetNames();
    }
    
    /**
     * Get the writable workbook that sheets will be added to.
     * @return workbook that will be written to.
     */
    public WritableWorkbook getResultBook() {
        return resultBook;
    }

    /**
     * Get current workbook that is being read from, can be null.
     *
     * @return Workbook being read from.
     */
    public Workbook getWorkbook() {
        return workBook;
    }

    /**
     * JOption message dialog shown on main GUI thread.
     *
     * @param msg String to display to user
     */
    public void displayMessage(String msg) {
        JOptionPane.showMessageDialog(null, msg);
    }

    /**
     * Print information to system out during debugging.
     *
     * @param msg Message to print
     */
    private void debug(String msg) {
        System.out.println(msg);
    }

    //Main method
    public static void main(String[] args) {
        SurveyProcessor sp = new SurveyProcessor();
        
        if (args.length != 2) {
            SurveyProcessorGUI gui = new SurveyProcessorGUI();
        } else {
            sp.loadWorkbook(args[0]);
            try {
                sp.processSheet(args[1]);
            } catch (WriteException | IOException ex) {
                Logger.getLogger(SurveyProcessor.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }
}
