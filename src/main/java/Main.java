import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Created by July on 18.10.2017.
 */
public class Main {
    public static void main(String[] args) {
        //Constructor of writer into exel
        IntoExel writer = new IntoExel();

        //Getting current date (for test)
        Date dNow = new Date( );
        SimpleDateFormat ft = new SimpleDateFormat ("E yyyy.MM.dd 'at' hh:mm:ss a zzz");
        System.out.println("Current Date: " + ft.format(dNow));

        //Calling the method that writes into .xls file
        try {
            writer.writeIntoExcel("resources/test.xls", "Zone 1", 30123, dNow, dNow, dNow);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
