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
    public static final short POSITIONS_ROW_END = 1000;
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
            try {
                order.setCarNumber( cell.getStringCellValue() );
            } catch ( IllegalStateException e ) {
                order.setCarNumber(
                        String.valueOf( cell.getNumericCellValue() ));
            }


            // E18 - order number
            cell = sheet.getRow( ORDER_NUMBER_ROW ).getCell( ORDER_NUMBER_CELL );
            try {
                order.setOrderNumber( cell.getStringCellValue() );
            } catch ( IllegalStateException e ) {
                order.setOrderNumber(
                        String.valueOf( cell.getNumericCellValue() ));
            }


            // Get positions (item properties) from excel
            int rowEnd = Math.max( POSITIONS_ROW_END, sheet.getLastRowNum() );
            for ( int rowNum = POSITIONS_ROW_START; rowNum < rowEnd; rowNum++ )
            {
                // Get row of the position
                Row row = sheet.getRow( rowNum );

                // Get the first cell
                cell = row.getCell( POSITIONS_CELL_ITEM_NAME );

                // There are no more item positions
                if( cell.getStringCellValue().isEmpty() )
                {
                    // Try find "TOTAL" row
                    for( int rowTotal = rowNum; rowTotal < rowEnd; rowTotal++ )
                    {
                        row = sheet.getRow( rowTotal );
                        cell = row.getCell( TOTAL_CELL_NAME );
                        if( cell.toString().contains( "TOTAL" ))
                        {
                            cell = row.getCell( TOTAL_CELL_BATTERIES );
                            try {
                                order.setBatteriesTotal( cell.getNumericCellValue() );
                            } catch ( IllegalStateException e ) {
                                order.setBatteriesTotal(
                                        Double.parseDouble( cell.getStringCellValue() ));
                            }

                            cell = row.getCell( TOTAL_CELL_PALLETS );
                            try {
                                order.setPalletsTotal( cell.getNumericCellValue() );
                            } catch ( IllegalStateException e ) {
                                order.setPalletsTotal(
                                        Double.parseDouble( cell.getStringCellValue() ));
                            }
                            break;
                        }
                    }
                    break;
                }

                Position position = new Position();

                // A23 - item name
                position.setItemName( cell.getStringCellValue() );


                // B23 - barcode
                cell = row.getCell( POSITIONS_CELL_ITEM_BARCODE );
                try {
                    position.setItemBarcode( cell.getStringCellValue() );
                } catch ( IllegalStateException e ) {
                    position.setItemBarcode(
                            String.valueOf( cell.getNumericCellValue() ));
                }


                // C23 - battery count
                cell = row.getCell( POSITIONS_CELL_BATTERIES_TOTAL );
                try {
                    position.setBatteryCount( cell.getNumericCellValue() );
                } catch ( IllegalStateException e ) {
                    position.setBatteryCount(
                            Double.parseDouble( cell.getStringCellValue() ));
                }

                // G23 - pallet count
                cell = row.getCell( POSITIONS_CELL_PALLET_TOTAL );
                try {
                    position.setPalletCount( cell.getNumericCellValue() );
                } catch ( IllegalStateException e )
                {
                    // Put 0 if cell has text
                    if( cell.toString().length() > 5 ) {
                        position.setPalletCount( 0 );
                    }
                    else {
                        position.setPalletCount(
                                Double.parseDouble( cell.getStringCellValue() ));
                    }
                }

                // Add item to position list of result DTO
                order.getPositions().add( position );
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
}
