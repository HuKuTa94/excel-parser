import model.Order;
import model.Position;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class Main
{
    // Input
    public static final String PATH_TO_FILE = "orders/Заказ.xlsx";

    // Excel cells
    public static final short CAR_NUMBER_ROW = 15;
    public static final short CAR_NUMBER_CELL = 4;
    public static final short ORDER_NUMBER_ROW = 17;
    public static final short ORDER_NUMBER_CELL = 4;

    // Position cells
    public static final short POSITIONS_ROW_START = 22;
    public static final short POSITIONS_CELL_ITEM_NAME = 0;
    public static final short POSITIONS_CELL_ITEM_BARCODE = 1;
    public static final short POSITIONS_CELL_BATTERIES_TOTAL = 2;
    public static final short POSITIONS_CELL_PALLET_TOTAL = 6;

    // Total cells
    public static final short TOTAL_CELL_NAME = 0;
    public static final short TOTAL_CELL_BATTERIES = 2;
    public static final short TOTAL_CELL_PALLETS = 6;


    public static void main( String[] args )
    {
        try ( FileInputStream fis = new FileInputStream( PATH_TO_FILE ))
        {
            XSSFWorkbook excelBook = new XSSFWorkbook( fis );
            XSSFSheet sheet = excelBook.getSheetAt( 0 );

            // Result data object of the order
            Order order = new Order();

            // E16 - car number
            Cell cell = sheet.getRow( CAR_NUMBER_ROW ).getCell( CAR_NUMBER_CELL );
            order.setCarNumber( getStringValueFromCell( cell ));

            // E18 - order number
            cell = sheet.getRow( ORDER_NUMBER_ROW ).getCell( ORDER_NUMBER_CELL );
            order.setOrderNumber( getStringValueFromCell( cell ));


            // Get positions (item properties) from excel
            int rowNum = POSITIONS_ROW_START;
            do {
                // Get row of the position
                Row row = sheet.getRow( rowNum );

                // Position DTO
                Position position = new Position();

                // A23 - item name
                cell = row.getCell( POSITIONS_CELL_ITEM_NAME );
                position.setItemName( cell.getStringCellValue() );

                // B23 - barcode
                cell = row.getCell( POSITIONS_CELL_ITEM_BARCODE );
                position.setItemBarcode( getStringValueFromCell( cell ));

                // C23 - battery count
                cell = row.getCell( POSITIONS_CELL_BATTERIES_TOTAL );
                position.setBatteryCount( getNumericValueFromCell( cell ));

                // G23 - pallet count
                cell = row.getCell( POSITIONS_CELL_PALLET_TOTAL );
                position.setPalletCount( getNumericValueFromCell( cell ));

                // Add item to position list of result DTO
                order.getPositions().add( position );

                rowNum++;
            }
            while ( !getStringValueFromCell( sheet.getRow( rowNum ).getCell( POSITIONS_CELL_ITEM_NAME )).isEmpty() );


            // Try find "TOTAL" row
            int rowEnd = sheet.getLastRowNum();
            for( int rowTotal = rowNum; rowTotal < rowEnd; rowTotal++ )
            {
                Row row = sheet.getRow( rowTotal );
                cell = row.getCell( TOTAL_CELL_NAME );

                if( cell.toString().contains( "TOTAL" ))
                {
                    cell = row.getCell( TOTAL_CELL_BATTERIES );
                    order.setBatteriesTotal( getNumericValueFromCell( cell ));

                    cell = row.getCell( TOTAL_CELL_PALLETS );
                    order.setPalletsTotal( getNumericValueFromCell( cell ));

                    break;
                }
            }
            System.out.println( order );
            excelBook.close();
        }
        catch ( FileNotFoundException e)
        {
            System.err.println( "File not found! The program will be closed!" );
        }
        catch ( IOException e )
        {
            System.err.println( "Impossible close the excel file! Maybe file used by other process!" );
        }
    }

    private static double getNumericValueFromCell( Cell cell ) {
        // Try to get numeric value from the cell
        try {
            return cell.getNumericCellValue();
        }

        // Cell has text type value
        catch ( IllegalStateException illegalStateException )
        {
            // Try parse number from text type (text contains a number)
            try {
                return Double.parseDouble( cell.getStringCellValue() );
            }

            // Text not contains a number, return zero
            catch ( NumberFormatException numberFormatException ) {
                return 0.0;
            }
        }
    }

    private static String getStringValueFromCell( Cell cell ) {
        // Try to get string value from the cell
        try {
            return cell.getStringCellValue();
        }
        // Cell has numeric type value
        catch ( IllegalStateException e ) {
            return String.valueOf( cell.getNumericCellValue() );
        }
    }
}
