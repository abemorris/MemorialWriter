package memorialwriter;

import java.io.File;
import java.io.IOException;
import javax.swing.JOptionPane;
import jxl.read.biff.BiffException;

public class MemorialWriter
{
  public static void main(String[] args)
    throws IOException, BiffException
  {
    String file = JOptionPane.showInputDialog("Enter the full file path (including the .xls on the end): ");
    ExcelReader er = new ExcelReader(new File(file));
    JOptionPane.showMessageDialog(null, "Text has been generated in the MemorialList.txt file.");
  }
}
