package model;

import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.ArrayList;
import java.util.List;

@NoArgsConstructor
public class Order
{
    // Cell E18
    @Getter
    @Setter
    private String orderNumber;

    // Cell E16
    @Getter
    @Setter
    private String carNumber;

    // Cell A23 and bellow
    @Getter
    @Setter
    private List<Position> positions = new ArrayList<>();

    // Batteries total cell
    @Getter
    @Setter
    private double batteriesTotal;

    // Pallets total cell
    @Getter
    @Setter
    private double palletsTotal;


    public boolean checkBatteriesTotal() {
        return batteriesTotal == getBatteriesSum();
    }

    public boolean checkPalletsTotal() {
        return palletsTotal == getPalletsSum();
    }

    private double getPalletsSum() {
        double result = 0;
        for( Position position : positions ) {
            result += position.getPalletCount();
        }
        return result;
    }

    private double getBatteriesSum() {
        double result = 0;
        for( Position position : positions ) {
            result += position.getBatteryCount();
        }
        return result;
    }

    @Override
    public String toString() {
        StringBuilder builder = new StringBuilder();
        builder.append( "Номер заказа: ").append( orderNumber ).append( "\n" )
                .append( "Номер машины: " ).append( carNumber ).append( "\n" )
                .append( "Позиции товаров:\nНаименование товара\t\tШтрих-код товара\t\tКол-во батарей\t\tКол-во паллет\n" );

        for( Position position : positions ) {
            builder.append( position.toString() ).append( "\n" );
        }

        builder.append( "ИТОГО:\nБатареи: ").append( batteriesTotal )
                .append( "; значения совпадают: " ).append( checkBatteriesTotal() )
                .append( "\nПаллеты: ").append( palletsTotal )
                .append( "; значения совпадают: ").append( checkPalletsTotal() );

        return builder.toString();
    }

}
