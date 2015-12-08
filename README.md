# Ken Burns Slideshow
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
	<li>Support displaying date for the images. Date value is parsed from the file name with format <em>yyyy-MM-dd</em>. If you want a date to be displayed, please rename the image files to the format for this to work. A valid example: <em>2015-12-02.jpg</em></li>
	<li>Support audio files to be played in loop at background. Please specify a folder path containing audio files in config.xml.</li>
	<li>Options to load large images at lower resolution to improve performance.</li>
</ul>
## Updates
<ul>
	<li>2015-12-08
	<ul>
		<li>Changed bitmap scaling option to high quality.</li>
	</ul>
	</li>
	<li>2015-12-07
	<ul>
		<li>Fixed images not being displayed fully.</li>
		<li>Added customizable option to fadeout for each image. Disabling fadeout might slightly improve performance.</li>
		<li>Added runtime options dialog what supports changing several settings in config.xml at runtime. The dialog is called by pressing F12.</li>
		<li>Other minor fixes.</li>
	</ul>
	</li>
	<li>2015-12-06
	<ul>
		<li>Added customizable option to show lower resolution for large images to improve performance.</li>
		<li>Added customizable option to lock movement direction to "down only" for images with height larger than screen height * 1.5. This setting suits better for portraits that usually have faces at the upper part of the image.</li>
		<li>Added customizable option to only show the upper / lower part (depending on movement direction) for images with height larger than screen height * 1.5. This is to avoid image moving too fast thus less elegant.</li>
		<li>Added customizable option to only show the left / right part (depending on movement direction) for images with width larger than screen width * 1.5. This is to avoid image moving too fast thus less elegant.</li>
		<li>Improved error handling when dealing with corruptions. Instead of showing error prompt and exit, program will now replace the problematic image with a black screen as a placeholder in case dates are being displayed.</li>
	</ul>
	</li>
</ul>

## Note
<ul>
	<li>The config.xml file serves as a configuration that is loaded at program start. Paths and other settings can be changed to local / relative path according to the location of the folders on your system.</li>
	<li>Press F12 at runtime for options.</li>
</ul>
