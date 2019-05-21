package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.matchers.AddressMatchers;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.*;

class AddressTest {
    @DisplayName("Should return a valid Address object according to the input data")
    @ParameterizedTest(name = "Input column:{0}, row:{1}, address:{2}, type:{3}, should lead to an address: {6} (C:{4}/R:{5}) with type {7}")
    @CsvSource({
            "2,3,,,2,3,C4,Default",
            "4,8,,FixedRow,4,8,E$9,FixedRow",
            "0,0,X22,,23,21,X22,Default",
            "0,0,C11,FixedColumn,2,10,$C11,FixedColumn",
            "0,0,A999,FixedRowAndColumn,0,998,$A$999,FixedRowAndColumn",
    })
    void getAddress(int col, int row, String addr, Cell.AddressType type, int expCol, int expRow, String expAddr, Cell.AddressType expType) {
        Address a = buildAddress(col, row, addr, type);
        assertThat(a,
                allOf(
                        AddressMatchers.hasRowNumber(expRow),
                        AddressMatchers.hasColumnNumber(expCol),
                        AddressMatchers.hasAddress(expAddr),
                        AddressMatchers.hasType(expType)
                )
        );
    }

    @DisplayName("Should return a valid column letter according to the input data")
    @ParameterizedTest(name = "Input column:{0}, row:{1}, address:{2}, type:{3}, should lead to the column letter: {4}")
    @CsvSource({
            "2,3,,,C",
            "4,8,,FixedRow,E",
            "0,0,X22,,X",
            "0,0,R21,FixedColumn,R",
            "0,0,AAB99,FixedRowAndColumn,AAB",
    })
    void getColumn(int col, int row, String addr, Cell.AddressType type, String expCol) {
        Address a = buildAddress(col, row, addr, type);
        assertThat(a,
                is(
                        AddressMatchers.hasColumnName(expCol)
                )
        );
    }

    @DisplayName("Should return true if (identical) input 1 and 2 are compared using the equals method")
    @ParameterizedTest(name = "Input 1 (C:{0}; R:{1}; A:{2}; T:{3}) compared with input 2: (C:{4}; R:{5}; A:{6}; T:{7}) should lead to {8}")
    @CsvSource({
            "0,0,A1,FixedColumn,0,0,A1,FixedColumn,true",
            "0,0,A1,FixedColumn,0,0,A1,FixedRow,true",
            "1,2,,Default,0,0,B3,Default,true",
            "1,2,,Default,0,0,B3,FixedRow,true",
            "0,0,X11,FixedColumn,0,0,X12,FixedColumn,false",
            "1,2,,Default,0,0,B4,Default,false",
            "1,2,,FixedRow,0,0,B4,Default,false",
    })
    void equalsTest(int col1, int row1, String addr1, Cell.AddressType type1, int col2, int row2, String addr2, Cell.AddressType type2, boolean expectedResult){
        Address a1 = buildAddress(col1, row1, addr1, type1);
        Address a2 = buildAddress(col2, row2, addr2, type2);
        assertThat(a1.equals(a2), is(expectedResult));
    }

    @DisplayName("Should return the expected cell address as string")
    @ParameterizedTest(name = "Address {0} (Type: {1}) should lead to a valid address string")
    @CsvSource({
            "A1,,A1",
            "A1,Default,A1",
            "B22,FixedRow,B$22",
            "AX1,FixedColumn,$AX1",
            "S775,FixedRowAndColumn,$S$775",
    })
    void toStringTest(String address, Cell.AddressType type, String expectedString){
        Address a = buildAddress(0,0,address, type);
        assertThat(a.toString(), is(expectedString));
    }

    @DisplayName("Should lead to an exception if an illegal address is defined")
    @ParameterizedTest(name = "Address C:{0}; R:{1}; A:{2}; T:{3} should lead to an exception")
    @CsvSource({
            "0,0,TTTTTTT1,",
            "0,0,$,Default",
            "0,0,125475,",
            "0,0,##15,",
            "-10,0,,",
            "12,-2,,Default",
            "-1,-1,,FixedRow",
            "99999999,22,,FixedColumn",
            "0,0,'',Default"
    })
    void invalidTest(int col, int row, String address, Cell.AddressType type){
        boolean isValid = true;
        try {
            Address a = buildAddress(col,row,address, type);
        }
        catch (Exception ex){
            isValid = false;
        }
        assertThat(isValid, equalTo(false));

    }


    public static Address buildAddress(int col, int row, String addr, Cell.AddressType type){
        if (addr == null){
            if (type == null){
                return new Address(col, row);
            }
            else {
                return new Address(col, row, type);
            }
        }
        else {
            if (type == null){
                return new Address(addr);
            }
            else {
                return new Address(addr, type);
            }
        }
    }


}
