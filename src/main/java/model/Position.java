package model;

import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@NoArgsConstructor
public class Position
{
    // Cell A23
    @Getter
    @Setter
    private String itemName;

    // Cell B23
    @Getter
    @Setter
    private String itemBarcode;

    // Cell C23

    @Getter
    @Setter
    private double batteryCount;

    // Cell G23
    @Getter
    @Setter
    private double palletCount;


    @Override
    public String toString() {
        StringBuilder builder = new StringBuilder();
        builder.append( itemName ).append( "\t\t")
                .append( itemBarcode ).append( "\t\t" )
                .append( batteryCount ).append( "\t\t" )
                .append( palletCount );

        return builder.toString();
    }
}
