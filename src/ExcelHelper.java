import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelHelper 
{
    private final int NO_KEYWORD_ERROR = -1;
    
    private Sheet sheet;
    private int num_cols;
    
    public ExcelHelper(String filename) throws IOException, InvalidFormatException
    {
        Workbook wb = WorkbookFactory.create(new File(filename));
        sheet = wb.getSheetAt(0);
        getSheetWidth();
        writeSentences();
    }
    // iterates through rows and forms sentences from the data
    private void writeSentences() throws FileNotFoundException, UnsupportedEncodingException
    {
        PrintWriter writer = new PrintWriter("output.txt", "UTF-8");
        
        int firstCol = getColumn("First");
        int lastCol = getColumn("Last");
        int noteCol = getColumn("Note");
        
        String firstName;
        String lastName;
        String note;
        
        String memorial;
        int row = 1; // start after the header row
        while(row <= sheet.getLastRowNum()) // check every row in the sheet
        {
            // if the row isn't empty
            if(sheet.getRow(row).getCell(0) != null && !sheet.getRow(row).getCell(0).getStringCellValue().equals(""))
            {
                // read rows from the spreadsheet
                firstName = sheet.getRow(row).getCell(firstCol).getStringCellValue();
                lastName = sheet.getRow(row).getCell(lastCol).getStringCellValue();
                note = sheet.getRow(row).getCell(noteCol).getStringCellValue();
                
                memorial = "In memory of" + splitNote(note) + " by " + firstName + " " + lastName + ".";
                memorial = memorial.replaceAll("\n", ""); // takes care of times when cell wrapping results in reading a newline
                
                System.out.println(memorial);
                writer.println(memorial);
            }
            row++;
        }        
        
        writer.close();
    }
    // splits up the note to find the name of the deceased
    private String splitNote(String note_data)
    {
        if(note_data.length() > 0)
        {
            int index1 = 0;
            int index2;
            while(!note_data.substring(0,index1).contains("of") && index1 < note_data.length()) // "of" generally precedes the deceased's name
            {
                index1++;
            }
            index2 = index1;
            while(note_data.charAt(index2) != '.' && index2 < note_data.length()) // go to the period or the end of the note cell.
            {
                index2++;
            }
            return note_data.substring(index1, index2) + ",";
        }
        else return "!!! NO NOTE !!!";
    }
    // returns the index of a column with a header that contains headerKeyWord
    private int getColumn(String headerKeyWord)
    {
        Row row = sheet.getRow(0);
        // cycles through all cells in first column to find header.
        for(int i = 0; i <= num_cols; i++) 
        {
            if(row.getCell(i).getStringCellValue().contains(headerKeyWord))
            {
                return i;
            }
        }
        return NO_KEYWORD_ERROR; // if this statement is reached than the header was not found
    }
    
    // helper method to find the length of the header row
    private void getSheetWidth()
    {
        Row row = sheet.getRow(0);
        num_cols = 0;
        while(!row.getCell(num_cols).getStringCellValue().equals(""))
        {
            num_cols++;
        }
    }    
}
