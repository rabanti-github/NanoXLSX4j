module ch.rabanti.nanoxlsx4j.core {
    requires java.xml;
    requires java.compiler;

    exports ch.rabanti.nanoxlsx4j.colors;
    exports ch.rabanti.nanoxlsx4j.enums;
    exports ch.rabanti.nanoxlsx4j.exceptions;
    exports ch.rabanti.nanoxlsx4j.styles;
    exports ch.rabanti.nanoxlsx4j.themes;
    exports ch.rabanti.nanoxlsx4j.utils;
    exports ch.rabanti.nanoxlsx4j.internal.interfaces;
    exports ch.rabanti.nanoxlsx4j;
    exports ch.rabanti.nanoxlsx4j.utils.internal.xml to
            ch.rabanti.nanoxlsx4j.reader,
            ch.rabanti.nanoxlsx4j.writer;
}
