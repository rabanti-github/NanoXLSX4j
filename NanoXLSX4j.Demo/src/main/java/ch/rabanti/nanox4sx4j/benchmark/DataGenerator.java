package ch.rabanti.nanox4sx4j.benchmark;

import com.github.javafaker.Faker;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Random;

public class DataGenerator {

    public enum DataType{
        Integers,
        Floats,
        Strings,
        Addresses,
        Dates,
        Booleans,
        Mixed,
    }

    private static List<Object> getSampleData(int seed, int numberOfEntries){
        Locale locale = Locale.US;
        Random random = new Random(seed);
        Random mixedRandom = new Random(seed);
        Faker faker = new Faker(locale, random);
        List<Object> output = new ArrayList<>(numberOfEntries);
        return null;
    }

    private static Object getSample(Faker instance, DataType type, Random mixedRandom){
        switch (type){

            case Integers:
                return instance.number().randomNumber();
            case Floats:
                return instance.number().randomDouble(16, Integer.MIN_VALUE, Integer.MIN_VALUE);
            case Strings:
                return instance.lorem().characters(0, 64, true, true);
            case Addresses:
                break;
            case Dates:
                break;
            case Booleans:
                break;
            case Mixed:
                break;
        }
        return null;
    }
}
