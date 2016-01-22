package memorialwriter;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

/**
 *
 * @author Abe
 */
public class ExcelReader
{
  ArrayList<String> dead;
  ArrayList<String> donors;
  int noteColumn = 0;
  int lastNameColumn = 0;
  int firstNameColumn = 0;
  Workbook data;
  Sheet sheet;
  
  public ExcelReader(File file) throws IOException, BiffException
  {
    this.dead = new ArrayList();
    this.donors = new ArrayList();
    
    this.data = Workbook.getWorkbook(file);
    this.sheet = this.data.getSheet(0);
    
    findColumns();
    populateArrays();
  }
  
  public void populateArrays() throws IOException
  {
    int count = 1;
    int numCells = this.sheet.getColumn(this.noteColumn).length;
    while (count < numCells)
    {
      int charCount = 0;
      int charCount2 = 0;
      String check = "";
      String temp = this.sheet.getCell(this.noteColumn, count).getContents();
      System.out.println(temp);
      while (temp.charAt(charCount) != '.') {
        charCount++;
      }
      while ((!check.contains("of ")) && (check.length() < charCount))
      {
        check = check + temp.charAt(charCount2);
        charCount2++;
      }
      String note = "";
      if (temp.contains("memory of")) {
        note = note + "In memory of ";
      } else if (temp.contains("honor of")) {
        note = note + "In honor of ";
      }
      temp = temp.substring(charCount2, charCount);
      this.dead.add(note + temp);
      System.out.print("Dead: " + temp);
      temp = this.sheet.getCell(this.firstNameColumn, count).getContents() + " " + this.sheet.getCell(this.lastNameColumn, count).getContents();
      System.out.println(" Donor: " + temp);
      this.donors.add(temp);
      count++;
    }
    compileDonors();
    FileWriter outFile = new FileWriter("MemorialList.rtf");
    int charCount2;

      for (int count2 = 0; count2 < this.dead.size(); count2++) {
        outFile.write((String)this.dead.get(count2) + ", by " + (String)this.donors.get(count2) + "." + "\n");
      }
      outFile.close();
  }
  public void findColumns()
  {
    int numColumns = this.sheet.getRow(1).length;
    for (int count = 0; count < numColumns; count++)
    {
      String contents = this.sheet.getCell(count, 0).getContents();
        if (contents.contains("Last")) {
            this.lastNameColumn = count;
        } 
        else if (contents.contains("First")) {
        this.firstNameColumn = count;
      } else if (contents.contains("Note")) {
        this.noteColumn = count;
      }
    }
  }
  
  
  public void compileDonors()
  {
    String[] deads = new String[this.dead.size()];
    String[] donor = new String[this.donors.size()];
    for (int count = 0; count < deads.length; count++)
    {
      deads[count] = ((String)this.dead.get(count));
      donor[count] = ((String)this.donors.get(count));
    }
    this.dead.clear();
    this.donors.clear();
    for (int count = 0; count < deads.length; count++) {
      for (int count2 = 0; count2 < deads.length; count2++) {
        if ((deads[count].equals(deads[count2])) && (count != count2))
        {
          int tmp124_123 = count; String[] tmp124_122 = donor;tmp124_122[tmp124_123] = (tmp124_122[tmp124_123] + ", " + donor[count2]);
          deads[count2] = "";
          donor[count2] = "";
        }
      }
    }
    for (int count = 0; count < deads.length; count++) {
      if (!deads[count].equals(""))
      {
        this.dead.add(deads[count]);
        this.donors.add(donor[count]);
      }
    }
  }
}