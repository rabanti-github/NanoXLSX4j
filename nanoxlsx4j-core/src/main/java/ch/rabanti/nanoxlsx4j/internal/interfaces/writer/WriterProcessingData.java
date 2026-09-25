/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.interfaces.writer;

import ch.rabanti.nanoxlsx4j.annotations.InternalApi;
import ch.rabanti.nanoxlsx4j.styles.StyleManager;
import ch.rabanti.nanoxlsx4j.styles.StyleRepository;

/**
 * Interface, used by writing processors (preparation)
 */
@InternalApi
public interface WriterProcessingData {

    /**
     * Gets the StyleManager instance, that can be accessed during the write preparation
     *
     * @return StyleManager instance
     */
    StyleManager getStyleManager();

    /**
     * Sets the StyleManager instance, that can be accessed during the write preparation
     *
     * @param styleManager StyleManager instance
     */
    void setStyleManager(StyleManager styleManager);

    /**
     * Gets the StyleRepository instance, that can be accessed during the write preparation
     *
     * @return StyleRepository instance
     */
    StyleRepository getStyleRepository();

    /**
     * Sets the StyleRepository instance, that can be accessed during the write preparation
     *
     * @param styleRepository StyleRepository instance
     */
    void setStyleRepository(StyleRepository styleRepository);
}
