.# Ken Burns Slideshow
Create instant & real-time full-screen slideshow with the Ken Burns effect for a set of images.
## Motivation
This idea came to me while watching a slideshow on iPad. How nice would it be if we can create instant slideshows on a Windows machine with only a single executable and a folder of images.
## Screenshot
<img class="alignnone size-full wp-image-8" src="http://carlchang.blog.com/files/2015/12/无标题.png" alt="" width="649" height="366" />
## Language &amp; Requirement
<ul>
	<li>WPF application written in VB.net</li>
	<li>.Net framework 4.5 or above.</li>
	<li>Config.xml file at application root.</li>
</ul>
## Features
<ul>
	<li>Create instant full-screen slideshow with an imitation of the Ken Burns effect for a set of folders that contains images. Images will be looped.</li>
	<li>Support displaying date for the images. Date value is parsed from the file name with format <em>yyyy-MM-dd</em>. Please rename the image files to the format for this to work. A valid example: <em>2015-12-02.jpg</em></li>
	<li>Support audio files to be played in loop at background. Please specify a folder path containing audio files in config.xml.</li>
</ul>
## Note
<ul>
	<li>The config.xml file serves as a configuration that is loaded at program start. Paths need to be changed to local / relative path according to the location of the folders on your system.</li>
</ul>
