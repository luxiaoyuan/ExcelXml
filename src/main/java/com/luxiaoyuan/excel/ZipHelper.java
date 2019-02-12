package com.luxiaoyuan.excel;


import java.io.File;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.Enumeration;
import java.util.zip.ZipEntry;

import org.apache.commons.compress.archivers.zip.ZipFile;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipTypes;
import org.apache.poi.openxml4j.opc.ZipPackage;
import org.apache.poi.openxml4j.opc.internal.ContentTypeManager;

public final class ZipHelper {

	/**
	 * Forward slash use to convert part name between OPC and zip item naming
	 * conventions.
	 */
	private final static String FORWARD_SLASH = "/";

	/**
	 * Buffer to read data from file. Use big buffer to improve performaces. the
	 * InputStream class is reading only 8192 bytes per read call (default value
	 * set by sun)
	 */
	public static final int READ_WRITE_FILE_BUFFER_SIZE = 8192;

	/**
	 * Prevent this class to be instancied.
	 */
	private ZipHelper() {
		// Do nothing
	}

	/**
	 * Retrieve the zip entry of the core properties part.
	 *
	 * @throws OpenXML4JException
	 *             Throws if internal error occurs.
	 */
	public static ZipEntry getCorePropertiesZipEntry(ZipPackage pkg) {
		PackageRelationship corePropsRel = pkg.getRelationshipsByType(
				PackageRelationshipTypes.CORE_PROPERTIES).getRelationship(0);

		if (corePropsRel == null)
			return null;

		return new ZipEntry(corePropsRel.getTargetURI().getPath());
	}

	/**
	 * Retrieve the Zip entry of the content types part.
	 */
	public static ZipEntry getContentTypeZipEntry(ZipPackage pkg) {
		Enumeration<? extends ZipEntry> entries = pkg.getZipArchive().getEntries();
		
		// Enumerate through the Zip entries until we find the one named
		// '[Content_Types].xml'.
		while (entries.hasMoreElements()) {
			ZipEntry entry = entries.nextElement();
			if (entry.getName().equals(
					ContentTypeManager.CONTENT_TYPES_PART_NAME))
				return entry;
		}
		return null;
	}

	/**
	 * Convert a zip name into an OPC name by adding a leading forward slash to
	 * the specified item name.
	 *
	 * @param zipItemName
	 *            Zip item name to convert.
	 * @return An OPC compliant name.
	 */
	public static String getOPCNameFromZipItemName(String zipItemName) {
		if (zipItemName == null)
			throw new IllegalArgumentException("zipItemName");
		if (zipItemName.startsWith(FORWARD_SLASH)) {
			return zipItemName;
		}
		return FORWARD_SLASH + zipItemName;
	}

	/**
	 * Convert an OPC item name into a zip item name by removing any leading
	 * forward slash if it exist.
	 *
	 * @param opcItemName
	 *            The OPC item name to convert.
	 * @return A zip item name without any leading slashes.
	 */
	public static String getZipItemNameFromOPCName(String opcItemName) {
		if (opcItemName == null)
			throw new IllegalArgumentException("opcItemName");

		String retVal = opcItemName;
		while (retVal.startsWith(FORWARD_SLASH))
			retVal = retVal.substring(1);
		return retVal;
	}

	/**
	 * Convert an OPC item name into a zip URI by removing any leading forward
	 * slash if it exist.
	 *
	 * @param opcItemName
	 *            The OPC item name to convert.
	 * @return A zip URI without any leading slashes.
	 */
	public static URI getZipURIFromOPCName(String opcItemName) {
		if (opcItemName == null)
			throw new IllegalArgumentException("opcItemName");

		String retVal = opcItemName;
		while (retVal.startsWith(FORWARD_SLASH))
			retVal = retVal.substring(1);
		try {
			return new URI(retVal);
		} catch (URISyntaxException e) {
			return null;
		}
	}

/**
 * Opens the specified file as a zip, or returns null if no such file exists
 *
 * @param file
 *            The file to open.
 * @return The zip archive freshly open.
 */
public static ZipFile openZipFile(File file) throws IOException {
   if (!file.exists()) {
      return null;
   }

   return new ZipFile(file);
}

	/**
	 * Retrieve and open a zip file with the specified path.
	 *
	 * @param path
	 *            The file path.
	 * @return The zip archive freshly open.
	 */
	public static ZipFile openZipFile(String path) throws IOException {
		File f = new File(path);

		if (!f.exists()) {
			return null;
		}

		return new ZipFile(f);
	}
}

