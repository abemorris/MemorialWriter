import java.io.File;
import java.io.IOException;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;


public class driver
{
    public static void main(String[] args) throws IOException, InvalidFormatException
    {
        // Setup for GUI objects
        JFileChooser choose = new JFileChooser();
        JFrame parent = new JFrame("Memorial Writer");
        parent.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        choose.setCurrentDirectory(new File(System.getProperty("user.home")));
        
        // Check return value of the file selector
        int result = choose.showOpenDialog(parent);
        if(result == JFileChooser.APPROVE_OPTION)
        {
            File data = choose.getSelectedFile();
            ExcelHelper eh = new ExcelHelper(data);
        }
        else  // exit without processing any data.
        {
            System.out.println("No File Selected");
        }
        
        parent.dispose(); // close the JFrame to end execution
    }
}