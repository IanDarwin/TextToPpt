package com.darwinsys.txt2ppt;

/** Enum of supported picture types.
 * Values chosen to coincide with list of published ints in XSLFPictureData
 * @author Ian Darwin
 */
public enum XSLFPictureType {
	UNUSED0("lose"), 
	UNUSED1("lose"), 
	EMF("emf"), // 2;
	WMF("wmf"), // 3;
	PICT("pict"), // 4;
	JPEG("jpg"), // 5;
	PNG("png"), // 6;
	DIB("dib"), // 7;
	GIF("gif"), // 8;
	TIFF("tif"), // 9;
	EPS("eps"), // 10;
	BMP("bmp"), // 11;
	WPG("wpg"), // 12;
	WDP("wdp"); // 13;
	
	private final String name;
	
	private XSLFPictureType(String name) {
		this.name = name;
	}

	public static XSLFPictureType valueOfFilename(String fileName) {
		fileName = fileName.toLowerCase();
		int last = fileName.lastIndexOf('.');
		String ext = fileName.substring(last + 1);
		for (XSLFPictureType type : values()) {
			if (type.name.equals(ext)) {
				return type;
			}
		}
		if (ext.equals("jpeg")) {
			return JPEG;
		}
		throw new IllegalArgumentException("Unmatched filetype in " + fileName);
	}
}
