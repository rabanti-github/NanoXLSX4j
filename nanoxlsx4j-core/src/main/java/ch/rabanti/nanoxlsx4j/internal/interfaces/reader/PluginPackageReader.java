/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.interfaces.reader;

import ch.rabanti.nanoxlsx4j.annotations.InternalApi;

/**
 * Defines a queued package reader that optionally reads one ZIP entry with a fixed, known path.
 *
 * <p>Remarks: <p>
 * The reader is executed once when its queue is processed. If {@link #getStreamEntryName()} contains a path,
 * that exact ZIP entry is opened and supplied to the reader. If the entry does not exist, the reader is skipped.
 * If the property is <see langword="null"/> or empty, the reader is executed with a <see langword="null"/> stream.
 * For example, returning {@code xl/custom.xml} requests that fixed package part.
 * </p>
 * <p>
 * Use {@link DiscoveryPackageReader} instead when the part path is obtained from OOXML relationships
 * and may contain counters or otherwise vary between packages, such as
 * {@code xl/externalLinks/externalLink1.xml} and {@code externalLink2.xml}.
 * </p>
 * </p>
 */
@InternalApi
public interface PluginPackageReader {

    /**
     * Gets the exact, case-sensitive path of the ZIP entry to read, or <see langword="null"/> to execute without an entry stream.
     * @return Stream entry name
     */
    String getStreamEntryName();
}
