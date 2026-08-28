package ch.rabanti.nanoxlsx4j.styles;

public class StyleRepository {

    private static StyleRepository instance;

    public static StyleRepository getInstance() {
        if (instance == null) {
            instance = new StyleRepository();
        }
        return instance;
    }

    public Style addStyle(Style style) {
        // DUMMY -> Impelemnt
        return null;
    }

}
