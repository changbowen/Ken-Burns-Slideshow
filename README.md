# Ken Burns Slideshow
Create instant & real-time full-screen slideshow with the Ken Burns effect for a set of images.
## Motivation
This idea came to me while watching a slideshow on iPad. How nice would it be if we can create instant slideshows on a Windows machine with only a single executable and a folder of images.
## Preview
- <img class="alignnone size-full wp-image-8" src="http://carlchang.blog.com/files/2015/12/无标题.png" alt="" width="649" height="366" /><img class="alignnone size-full wp-image-20" src="http://carlchang.blog.com/files/2015/12/options.png" alt="" width="422" height="395" />
- Youtube video
  - https://youtu.be/zyt4tWciIdQ
- Youku video (for China)
  - http://v.youku.com/v_show/id_XMTQzODcxNDc4MA==.html

## Language &amp; Requirement
- WPF application written in VB.net
- .Net framework 4.5 or above.
- Config.xml file at application root.

## Features
- Create instant full-screen slideshow with an imitation of the Ken Burns effect for a set of folders that contains images. Images will be looped.
- Support displaying date for the images. Date value is parsed from the file name with format <em>yyyy-MM-dd</em>. If you want a date to be displayed, please rename the image files to the format for this to work. A valid example: <em>2015-12-02.jpg</em>
- Support audio files to be played in loop at background.
- Options to load large images at lower resolution to improve performance.

## Updates
- 2016-01-16
  - **Brand-new animation "Breath" added!**
  - Fixed fadeout setting not getting saved.
- 2016-01-15
  - Fixed crash when loading the first image fails.
  - Added icon which doubled the footprint of the executable.
- 2016-01-13
  - Added complete UI for modification of config.xml file. Removed config.xml from release as it is no longer required for program to start correctly.
  - Added some customizable parameters.
  - Other non-performance related tweaks.
- 2016-01-08
  - Fixed animation not being displayed correctly when last few dates share the same month value.
  - Fixed animation not being displayed correctly when vertical lock is on.
- 2016-01-06
  - Now mouse cursor will be hidden while slideshow is playing.
- 2015-12-08
  - Changed bitmap scaling option to high quality.
- 2015-12-07
  - Fixed images not being displayed fully.
  - Added customizable option to fadeout for each image. Disabling fadeout might slightly improve performance.
  - Added runtime options dialog what supports changing several settings in config.xml at runtime. The dialog is called by pressing F12.
  - Other minor fixes.
- 2015-12-06
  - Added customizable option to show lower resolution for large images to improve performance.
  - Added customizable option to lock movement direction to "down only" for images with height larger than screen height * 1.5. This setting suits better for portraits that usually have faces at the upper part of the image.
  - Added customizable option to only show the upper / lower part (depending on movement direction) for images with height larger than screen height * 1.5. This is to avoid image moving too fast thus less elegant.
  - Added customizable option to only show the left / right part (depending on movement direction) for images with width larger than screen width * 1.5. This is to avoid image moving too fast thus less elegant.
  - Improved error handling when dealing with corruptions. Instead of showing error prompt and exit, program will now replace the problematic image with a black screen as a placeholder in case dates are being displayed.

## Upcoming
- Working on more slide animations.

## Note
- ~~The config.xml file serves as a configuration that is loaded at program start. Paths and other settings can be changed to local / relative path according to the location of the folders on your system.~~
- Press F12 at runtime for options.
- ~~Please enclose the folder path with `<![CDATA[     ]]>` when the path contains markup chars in XML like &.~~
  - ~~Example: `<![CDATA[F:\Folder with & in name\images]]>`~~
- Animation might not be as smooth when not using Windows Aero themes on Windows 7.
