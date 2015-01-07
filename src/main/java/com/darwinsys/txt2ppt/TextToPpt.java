package com.darwinsys.txt2ppt;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

/**
 * Import a tab-indented text file into a PowerPoint 2010 pptx file
 * Sample input:
 * <pre>
 * My Slide Show
 * My Slide #1 Title
 * 	Bullet #1
 * 	Bullet #2
 * 	Bullet #3
 * </pre>
 * @author Ian Darwin
 * @author A lot of the POI API code cribbed from Apache POI example code at
 * http://poi.apache.org/slideshow/xslf-cookbook.html
 */
public class TextToPpt {

	/** Main program
	 * @param args One or more input filenames; if none, a built-in demo file is processed.
	 */
	public static void main(String[] args) {
		final TextToPpt program = new TextToPpt();
		BufferedReader is;
		try {
			if (args.length > 0) {
				for (String arg : args) {
					is = getReaderFor(arg);
					program.readAndProcess(is);
					is.close();
				}
			} else {
				InputStream bis = program.getClass().getResourceAsStream("/demoshow.txt");
				is = new BufferedReader(new InputStreamReader(bis));
				saveShow(program.readAndProcess(is), "/tmp/demoshow.pptx");
				is.close();
			}
		} catch (IOException e) {
			System.err.println("Unexpected exception " + e);
		}
	}
	
	XMLSlideShow template;
	XSLFSlideMaster defaultMaster;

	/**
	 * Read one text file in the format above, and output a PPTX file
	 * @param is A BufferedReader for inputting
	 * @return The complete(?) XMLSlideShow generated from the input.
	 */
	private XMLSlideShow readAndProcess(BufferedReader is) {
		readTemplate(getInputStreamFor("template/ltree/Ch00 2012.potx"), "POTX");
		XMLSlideShow show = createShow();
		List<Item> items = new ArrayList<>();
		try {
			String line = is.readLine();
			doChapterTitleSlide(show, line);
			// XXX Need to implement post-handling here!
			while ((line = is.readLine()) != null) {
				System.out.println("Input line: " + line);
				if (line.startsWith("\t")) {
					// XXX save it
				} else {
					addSlide(show, line, items);
				}
			}
		} catch (IOException e) {
			throw new RuntimeException("Could not read file input");
		}
		return show;
	}
	
	/**
	 * Save a new show to disk
	 * @param show THe XMLSlideShow object
	 * @param fileName The output file to (over)write
	 * @throws IOException If anything goes wrong!
	 */
	static void saveShow(XMLSlideShow show, String fileName) throws IOException {
		try (FileOutputStream out = new FileOutputStream(fileName)) {
			show.write(out);
    	}
	}

	/**
	 * Read one "POTX"-style template file into memory
	 * @param is The inputstream opened to the file.
	 * @param templateFileName The filename, only for use in messages
	 */
	private void readTemplate(InputStream is, String templateFileName) {
		try {
			template = new XMLSlideShow(is);

			// first see what slide layouts are available :
			System.out.println("Available slide layouts:");
			for (XSLFSlideMaster master : template.getSlideMasters()){
				for(XSLFSlideLayout layout : master.getSlideLayouts()){
					System.out.println(layout.getType());
				}
			}

			// there can be multiple masters each referencing a number of layouts
			// for demonstration purposes we use the first (default) slide master
			defaultMaster = template.getSlideMasters()[0];

		} catch (IOException ex) {
			throw new IllegalArgumentException("Can't open " + templateFileName);
		}
	}

	XMLSlideShow createShow() {
		//create a new empty slide show
		return new XMLSlideShow();
	}

	/**
	 * Creates the title slide, from the first line of the text file.
	 * @param show The already-created XMLSlideShow
	 * @param title The string to place on the title
	 */
	private void doChapterTitleSlide(XMLSlideShow show, String title) {
		System.out.println("TextToPpt.doChapterTitleSlide()");
		// there can be multiple masters each referencing a number of layouts
		// for demonstration purposes we use the first (default) slide master
		XSLFSlideMaster defaultMaster = show.getSlideMasters()[0];
		// title slide
		XSLFSlideLayout titleLayout = defaultMaster.getLayout(SlideLayout.TITLE);
		// fill the placeholders
		XSLFSlide slide1 = show.createSlide(titleLayout);
		XSLFTextShape title1 = slide1.getPlaceholder(0);
		title1.setText(title);
	}
	
	private XSLFSlide addSlide(XMLSlideShow show, String slideTitle, List<Item> body) {
		// first see what slide layouts are available
		System.out.println("TextToPpt.createSlide(): Available slide layouts:");
		for (XSLFSlideMaster master : show.getSlideMasters()){
			for (XSLFSlideLayout layout : master.getSlideLayouts()){
				System.out.println(layout.getType());
			}
		}

		// title and content
		XSLFSlideLayout titleBodyLayout = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
		XSLFSlide slide = show.createSlide(titleBodyLayout);

		XSLFTextShape title = slide.getPlaceholder(0);
		title.setText(slideTitle);

		XSLFTextShape body2 = slide.getPlaceholder(1);
		body2.clearText(); // unset any existing text
		body2.addNewTextParagraph().addNewTextRun().setText("First paragraph");
		body2.addNewTextParagraph().addNewTextRun().setText("Second paragraph");
		body2.addNewTextParagraph().addNewTextRun().setText("Third paragraph");

		return slide;
	}

	XSLFPictureShape addImage(XMLSlideShow show, XSLFSlide slide, String fileName) {
		byte[] pictureData;
		try {
			pictureData = IOUtils.toByteArray(new FileInputStream(fileName));
		} catch (IOException e) {
			throw new RuntimeException("Can't read input file " + fileName, e);
		}

        int idx = show.addPicture(pictureData, XSLFPictureData.PICTURE_TYPE_PNG);
        return slide.createPicture(idx);
	}
	
	static String[] paths = {
		"src/main/resources",
		System.getProperty("user.home")	
	};
	static InputStream getInputStreamFor(String fileName) {
		File f = new File(fileName);
		try {
			if (f.exists()) {
					return new FileInputStream(f);
			}
			for (String d : paths) {
				if ((f = new File(d, fileName)).exists()) {
					return new FileInputStream(f);
				}
			}
		} catch (FileNotFoundException e) {
			throw new IllegalArgumentException("Unexpected error opening " + f);
		}
		throw new IllegalArgumentException("Looked all over, can't find " + fileName);
	}
	
	static BufferedReader getReaderFor(String fileName) {
		File f = new File(fileName);
		try {
			if (f.exists()) {
					return new BufferedReader(new FileReader(f));
			}
			for (String d : paths) {
				if ((f = new File(d, fileName)).exists()) {
					return new BufferedReader(new FileReader(f));
				}
			}
		} catch (FileNotFoundException e) {
			throw new IllegalArgumentException("Unexpected error opening " + f);
		}
		throw new IllegalArgumentException("Looked all over, can't find " + fileName);
	}
}
