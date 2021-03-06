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

import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFNotes;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
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
	private static final String OUTPUT_EXTENSION = ".pptx";
	private static final String DEFAULT_TEMPLATE = "Chapter 2014.pptx";
	static int fileNumber;
	private static final boolean verbose = false;

	/** Main program
	 * @param args One or more input filenames; if none, a built-in demo file is processed.
	 */
	public static void main(String[] args) {
		String template = DEFAULT_TEMPLATE;
		final TextToPpt program = new TextToPpt();
		BufferedReader is;
		try {
			if (args.length > 0) {
				for (String arg : args) {
					is = getReaderFor(arg);
					saveShow(program.readAndProcessOneFile(is, template), generateFileName(arg));
					is.close();
				}
			} else {
				InputStream bis = program.getClass().getResourceAsStream("/demoshow.txt");
				if (bis == null) {
					System.err.println("Usage: " + TextToPpt.class.getSimpleName() + " textFile [...]");
					System.exit(1);
				}
				is = new BufferedReader(new InputStreamReader(bis));
				saveShow(program.readAndProcessOneFile(is, template), generateFileName("/tmp/demoshow.txt"));
				is.close();
			}
		} catch (IOException e) {
			System.err.println("Unexpected exception " + e);
		}
	}
	
	/** Output filename based on input name, but replace ".txt" or whatever with ".pptx" */
	private static String generateFileName(String inputFileName) {
		int dot = inputFileName.lastIndexOf('.');
		if (dot == -1) {
			return String.format("/tmp/generated%d.pptx");
		}
		return inputFileName.substring(0, dot) + OUTPUT_EXTENSION;
	}
	
	XMLSlideShow show;
	XSLFSlideMaster defaultMaster;

	/**
	 * Read one text file in the format above, and output a PPTX file
	 * @param is A BufferedReader for inputting
	 * @return The complete(?) XMLSlideShow generated from the input.
	 */
	private XMLSlideShow readAndProcessOneFile(BufferedReader is, String template) {
		show = readTemplate(getInputStreamFor(template), "POTX");
		boolean inCode = false;
		try {
			int lineNumber = 0;
			String line = is.readLine(); ++lineNumber;
			doChapterTitleSlide(show, line); // First line of file is chapter title
			
			XSLFSlide slide = null;
			XSLFTextShape body = null;
			int thisIndent = 0, codeIndent = 0;
			
			// MAIN LOOP
			while ((line = is.readLine()) != null) {
				++lineNumber;
				String trimmedLine = line.trim();
				if (trimmedLine.length() == 0) {
					continue;
				}
				
				// Strip comment lines, but only at the left margin - otherwise,
				// what if the slide contains a scripting language example?
				if (line.charAt(0) == '#') {
					continue;
				}
				
				// Code insert? Syntax stolen from AsciiDoc
				if (trimmedLine.equals("----")) {
					inCode = !inCode;
					if (inCode) {
						codeIndent = thisIndent;
					}
					continue;
				}
				
				if (line.charAt(0) == ' ') {
					System.err.println("Warning: Leading spaces on line " +
						lineNumber + ", trying to correct to tabs.");
					line = line.replaceAll("    ", "\t");
				}
				if (verbose) {
					System.out.println("Input line " + lineNumber + ": " + line);
				}
				thisIndent = 0;
				while (line.charAt(thisIndent) == '\t') {
							++thisIndent;
				}
				
				// An Image?
				if (trimmedLine.startsWith("IMAGE")) {
					String fileName = trimmedLine.substring(5).trim();
					System.out.println("IMAGE at line " + lineNumber + ": " + fileName);
					XSLFPictureType type = null;
					try {
						type = XSLFPictureType.valueOfFilename(fileName);
					} catch (Exception e) {
						System.err.println("Unknown image type " + fileName);
						continue;
					}
					byte[] pictureData = IOUtils.toByteArray(new FileInputStream(fileName));
			        int idx = show.addPicture(pictureData, type.ordinal());
			        slide.createPicture(idx);
			        continue;
				}
				
				// Speaker Note?
				if (trimmedLine.startsWith("NOTE")) {
					String noteText = trimmedLine.substring(4).trim();
					System.out.println("NOTE at line " + lineNumber + ": " + noteText);
					final XSLFNotes notes = show.getNotesSlide(slide);
					//notes.getPlaceholder(0).setText(noteText);
					final XSLFTextBox textBox = notes.createTextBox();
					final XSLFTextParagraph notesPara = textBox.addNewTextParagraph();
					notesPara.addNewTextRun().setText(noteText);
					continue;
				}
				
				if (thisIndent == 0) {
					// First line with no tabs is next title, so start new slide
					System.out.println("TextToPpt.createSlide()");

					if (inCode) {
						System.out.println("WARNING: code block not closed by line " + lineNumber);
						inCode = false;
					}

					// title and content
					XSLFSlideLayout titleBodyLayout = defaultMaster.getLayout(SlideLayout.CUST);
					slide = show.createSlide(titleBodyLayout);

					XSLFTextShape title1 = slide.getPlaceholder(0);
					title1.setText(line);

					body = slide.getPlaceholder(1);
					body.clearText(); // unset any existing text
					continue;
				}
				
				// Else a regular line
				final XSLFTextParagraph para = body.addNewTextParagraph();
				final XSLFTextRun run = para.addNewTextRun();
				run.setText(line.substring(thisIndent));
				if (inCode) {
					para.setBullet(false);
					para.setLevel(thisIndent - 1);
					run.setFontFamily("Courier New");
				} else {
					para.setLevel(thisIndent - 1);
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
			System.out.println("Saved show to " + fileName);
    	}
	}

	/**
	 * Read one "POTX"-style template file into memory
	 * @param is The inputstream opened to the file.
	 * @param templateFileName The filename, only for use in messages
	 * @return The slide show represented by the template file.
	 */
	private XMLSlideShow readTemplate(InputStream is, String templateFileName) {
		try {
			XMLSlideShow template = new XMLSlideShow(is);

			// first see what slide layouts are available :
			System.out.println("Available slide layouts:");
			for (XSLFSlideMaster master : template.getSlideMasters()){
				for(XSLFSlideLayout layout : master.getSlideLayouts()){
					System.out.println(layout.getType());
				}
			}

			// There can be multiple masters each referencing a number of layouts.
			// For demonstration purposes we use the first (default) slide master
			defaultMaster = template.getSlideMasters()[0];
			
			// The template may have slides, which we'll get rid of here.
			// Except that that crashes with this POI error:
			// PartAlreadyExistsException: A part with the name '/ppt/slides/slide3.xml' already exists...
			// for (int i = 0; i < template.getSlides().length; i++) {
			//	template.removeSlide(i);
			//}

			return template;
		} catch (IOException ex) {
			throw new IllegalArgumentException("Can't open " + templateFileName);
		}
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
		XSLFSlideLayout titleLayout = defaultMaster.getLayout(SlideLayout.CUST);
		// fill the placeholders
		XSLFSlide slide1 = show.createSlide(titleLayout);
		XSLFTextShape title1 = slide1.getPlaceholder(0);
		title1.setText(title);
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
		System.getProperty("user.home") + "/template/ltree",
		System.getProperty("user.home")	
	};
	static InputStream getInputStreamFor(String fileName) {
		File f = new File(fileName);
		try {
			if (f.exists()) {
					return new FileInputStream(f);
			}
			for (String d : paths) {
				f = new File(d, fileName);
				if (verbose) {
					System.out.printf("TextToPpt.getInputStreamFor(%s): try %s%n",
						fileName, f);
				}
				if (f.exists()) {
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
				f = new File(d, fileName);
				System.out.printf("TextToPpt.getInputStreamFor(%s): try %s%n",
						fileName, f);
				if (f.exists()) {
					return new BufferedReader(new FileReader(f));
				}
			}
		} catch (FileNotFoundException e) {
			throw new IllegalArgumentException("Unexpected error opening " + f);
		}
		throw new IllegalArgumentException("Looked all over, can't find " + fileName);
	}
}
