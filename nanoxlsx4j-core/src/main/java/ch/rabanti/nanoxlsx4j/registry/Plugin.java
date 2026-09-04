/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.registry;

/**
 * Marker interface for NanoXLSX4j service providers.
 * <p>
 * Reader- and writer-specific plug-in interfaces extend this interface when those contracts are available.
 * Implementations are discovered through {@link java.util.ServiceLoader}.
 */
public interface Plugin {
}
