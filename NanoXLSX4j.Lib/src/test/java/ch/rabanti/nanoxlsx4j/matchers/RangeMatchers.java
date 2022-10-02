package ch.rabanti.nanoxlsx4j.matchers;

import static org.hamcrest.Matchers.equalTo;

import org.hamcrest.FeatureMatcher;
import org.hamcrest.Matcher;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Range;

public class RangeMatchers {

	public static Matcher<Range> hasStartAddress(Matcher<? super Address> matcher) {
		return new FeatureMatcher<Range, Address>(matcher, "resolved start address", "startAddress") {
			@Override
			protected Address featureValueOf(Range range) {
				return range.StartAddress;
			}
		};
	}

	public static Matcher<Range> hasStartAddress(Address startAddress) {
		return hasStartAddress(equalTo(startAddress));
	}

	public static Matcher<Range> hasEndAddress(Matcher<? super Address> matcher) {
		return new FeatureMatcher<Range, Address>(matcher, "resolved end address", "endAddress") {
			@Override
			protected Address featureValueOf(Range range) {
				return range.EndAddress;
			}
		};
	}

	public static Matcher<Range> hasEndAddress(Address endAddress) {
		return hasEndAddress(equalTo(endAddress));
	}

	public static Matcher<Range> hasRangeString(Matcher<? super String> matcher) {
		return new FeatureMatcher<Range, String>(matcher, "resolved range string", "rangeString") {
			@Override
			protected String featureValueOf(Range range) {
				return range.toString();
			}
		};
	}

	public static Matcher<Range> hasRangeString(String rangeString) {
		return hasRangeString(equalTo(rangeString));
	}
}
