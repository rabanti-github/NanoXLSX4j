package ch.rabanti.nanoxlsx4j.matchers;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Cell;
import org.hamcrest.FeatureMatcher;
import org.hamcrest.Matcher;

import static org.hamcrest.Matchers.equalTo;

public class AddressMatchers {
    public static Matcher<Address> hasRowNumber(Matcher<? super Integer> matcher) {
        return new FeatureMatcher<Address, Integer>(matcher, "zero-based row number",
                                                    "row"
        ) {
            @Override
            protected Integer featureValueOf(Address address) {
                return address.Row;
            }
        };
    }

    public static Matcher<Address> hasRowNumber(Integer row) {
        return hasRowNumber(equalTo(row));
    }

    public static Matcher<Address> hasColumnNumber(Matcher<? super Integer> matcher) {
        return new FeatureMatcher<Address, Integer>(matcher, "zero-based column number",
                                                    "column"
        ) {
            @Override
            protected Integer featureValueOf(Address address) {
                return address.Column;
            }
        };
    }

    public static Matcher<Address> hasColumnNumber(Integer column) {
        return hasColumnNumber(equalTo(column));
    }

    public static Matcher<Address> hasType(Matcher<? super Cell.AddressType> matcher) {
        return new FeatureMatcher<Address, Cell.AddressType>(matcher, "address type",
                                                             "type"
        ) {
            @Override
            protected Cell.AddressType featureValueOf(Address address) {
                return address.Type;
            }
        };
    }

    public static Matcher<Address> hasType(Cell.AddressType type) {
        return hasType(equalTo(type));
    }

    public static Matcher<Address> hasAddress(Matcher<? super String> matcher) {
        return new FeatureMatcher<Address, String>(matcher, "resolved address",
                                                   "address"
        ) {
            @Override
            protected String featureValueOf(Address address) {
                return address.getAddress();
            }
        };
    }

    public static Matcher<Address> hasAddress(String address) {
        return hasAddress(equalTo(address));
    }

    public static Matcher<Address> hasColumnName(Matcher<? super String> matcher) {
        return new FeatureMatcher<Address, String>(matcher, "resolved column name",
                                                   "columnName"
        ) {
            @Override
            protected String featureValueOf(Address address) {
                return address.getColumn();
            }
        };
    }

    public static Matcher<Address> hasColumnName(String colName) {
        return hasColumnName(equalTo(colName));
    }
}
