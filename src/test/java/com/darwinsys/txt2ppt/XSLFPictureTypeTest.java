package com.darwinsys.txt2ppt;

import static org.junit.Assert.assertEquals;

import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.junit.Test;

public class XSLFPictureTypeTest {

	@Test
	public void testPng() {
		assertEquals("PNG", XSLFPictureType.PNG, XSLFPictureType.valueOfFilename("foo.png"));
	}
	
	@Test
	public void testJpeg() {
		assertEquals("JPEG", XSLFPictureType.JPEG, XSLFPictureType.valueOfFilename("foo.jpg"));
		assertEquals("JPEG", XSLFPictureType.JPEG, XSLFPictureType.valueOfFilename("foo.jpeg"));
	}

	@Test
	public void testOrdinals() {
		assertEquals("WMF", XSLFPictureData.PICTURE_TYPE_WMF, XSLFPictureType.WMF.ordinal());
	}
}
