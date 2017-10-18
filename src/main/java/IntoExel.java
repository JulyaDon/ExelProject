import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

/**
 * Created by July on 18.10.2017.
 * Implements writing into .xls file
 */
public class IntoExel {
    /**
     * Writes data into .xls file
     * Creates a list with data from exact zone
     * @param file - path to file into which we write
     * @param zone - zone with objects
     * @param id - id of object
     * @param timeIn - time at which the object entered the zone
     * @param timeOut - time at which the leaved the zone
     * @param timeInZone - quantity of time the object was in zone
     * @throws FileNotFoundException
     * @throws IOException
     */
    @SuppressWarnings("deprecation")
    public static void writeIntoExcel(String file, String zone, long id, Date timeIn, Date timeOut, Date timeInZone) throws FileNotFoundException, IOException {
        Workbook book = new HSSFWorkbook();
        Sheet sheet = book.createSheet(zone);

        //First row with names of columns
        Row row1 = sheet.createRow(0);

        Cell idCell = row1.createCell(0);
        idCell.setCellValue("ID");
        Cell timeInCell = row1.createCell(1);
        timeInCell.setCellValue("Time in");
        Cell timeOutCell = row1.createCell(2);
        timeOutCell.setCellValue("Time out");
        Cell timeInZoneCell = row1.createCell(3);
        timeInZoneCell.setCellValue("Time in zone");

        //Second row with data
        Row row2 = sheet.createRow(1);

        //ID of object
        Cell id1 = row2.createCell(0);
        id1.setCellValue(Long.toString(id));

        //Time the object entered the zone
        Cell timeIn1 = row2.createCell(1);
        DataFormat format = book.createDataFormat();
        CellStyle dateStyle = book.createCellStyle();
        dateStyle.setDataFormat(format.getFormat("hh:mm:ss"));
        timeIn1.setCellStyle(dateStyle);
        timeIn1.setCellValue(timeIn);

        //Time the object exited the zone
        Cell timeOut1 = row2.createCell(2);
        dateStyle.setDataFormat(format.getFormat("hh:mm:ss"));
        timeOut1.setCellStyle(dateStyle);
        timeOut1.setCellValue(timeOut);

        //Time the object spent in the zone
        Cell timeInZone1 = row2.createCell(3);
        dateStyle.setDataFormat(format.getFormat("hh:mm:ss"));
        timeInZone1.setCellStyle(dateStyle);
        timeInZone1.setCellValue(timeInZone);

        //Making the autosize of column because ID can be very big
        sheet.autoSizeColumn(1);

        //Writing our data into file
        book.write(new FileOutputStream(file));
        book.close();
    }
}
