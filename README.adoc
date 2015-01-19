== TextToPpt

Imports a tab-indented text file into a PowerPoint 201x pptx file.
Each tab indent will mean an indentation level (up to 4).

=== Usage

Assuming you have ~/lib for your jars and have done _mvn package_ and run _inst_

java -jar ~/lib/textToPpt.jar *.txt

=== Sample input:
 
----
My Superb Slide Show
My Slide #1 Title
	Bullet #1
		SubBullet #1
		SubBullet #2
	Bullet #2
	Bullet #3
My Slide #2
	This is how it's done:
	----
	public class Foo {
		//
	}
	----
	That was a code insert.
My Troisieme Slide
	This is what it should look like:
	IMAGE images/screenshot.png
----

=== BUGS/TODO

* Needs parameterization ("getopt" or config file): has some dependencies on my personal setup and workflow.
* Need a syntax to specify which layout (hardcoding CUST isn't great; should allow either a number or
a SlideLayout enum label (pick with SlideLayout.valueOf(verb.toUpperCase())
* IMAGE takes any file but only formats supported by XSLFPictureData type work (and this does not include PDF). Image placement needs manual intervention.
* Need a syntax to insert speaker notes ("NOTE" as first word)
* You can use a .potx or a .pptx (or older) for your template, but any slides in it will not be removed.
* There's still room to improve the handling of leading tabs in code.
* This may all be obviated when AsciiDoctor gets the ability to output slide shows fully working.

=== About
Author Ian Darwin. 

License: ASL2 because of lines derived from Apache POI code examples.
