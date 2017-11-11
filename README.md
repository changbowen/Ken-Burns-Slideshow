# Ken Burns Slideshow
Instant, portable, full-screen slideshow application with the Ken Burns effect (**and more**)!
## Motivation
This idea came to me while watching a slideshow on iPad. How nice would it be if we can create instant slideshows on a Windows machine with only a single executable and a folder of images.

此程序可用Ken Burns等三种效果播放文件夹中的图像，Ken Burns效果可参照iPad照片图库。

Always get the latest development build from [here](https://github.com/changbowen/Ken-Burns-Slideshow/raw/master/bin/Release/Ken%20Burns%20Slideshow.exe).

最新开发版[传送门](https://github.com/changbowen/Ken-Burns-Slideshow/raw/master/bin/Release/Ken%20Burns%20Slideshow.exe)。

If you really like it [![PayPal](https://img.shields.io/badge/%24-PayPal-blue.svg)](https://www.paypal.me/BowenChang) or [支付宝](https://user-images.githubusercontent.com/15975872/29361889-175fef58-82bc-11e7-9e3b-ed3c748456b8.png)

------------
### I want to rewrite this thing with C# in the future when I have more free time (when will that be...). The current VB project was written a long time ago and it consists of poorly organized codes with little descriptive comments that makes it so hard to continue to work on...
### 我有计划想用C#重写这个程序（当我有时间的时候。。）。现在的VB代码是很久之前写的，很乱并且缺少说明，很难继续改良。。
------------

## Preview
- <img src="http://i.imgur.com/nbznvOh.gif" title="Ken Burns effect preview"/>
- <img src="http://i.imgur.com/A97UmCm.gif" title="Breath effect preview"/>
- <img src="http://i.imgur.com/d7Ap7t5.gif" title="Throw effect preview"/>
- <img src="http://i.imgur.com/5O30vFL.jpg" width="600" title="Supports date & custom text display"/>
- <img src="http://i.imgur.com/dNIf5mC.jpg" width="600" title="Detailed customizations"/>
- [Youtube video](https://youtu.be/ch2UjN9nwIc)
- [Youku video](http://v.youku.com/v_show/id_XMTQ5NTM0NTAxMg==.html) 

## Language &amp; Requirement
- WPF application written in VB.net
- .Net framework 4.5 or above.
- ~~Config.xml file at application root.~~
- 此程序基于WPF，使用VB.net编写。需要.Net framework 4.5或以上。

## Features
- Create instant full-screen slideshow with an imitation of the Ken Burns effect (**and more!**) for a set of folders that contains images. Images will be looped.
- Support displaying date for the images. Date value is parsed from the file name with format <em>yyyy-MM-dd</em>. If you want a date to be displayed, please rename the image files to the format for this to work. A valid example: <em>2015-12-02.jpg</em>
- Support audio files to be played in loop at background.
- Options to load large images at lower resolution to improve performance.
- Options for each slide (Edit Mode) is now available by pressing F11.
- Support launching slideshow directly from a folder containing images and audio files.
- Support EXIF rotation.
- Simplified Chinese language support.

## 功能特性
- 首次使用需要设置图片文件夹和音乐文件夹（可选）。
- 支持在每张图片上面显示日期。如果图片文件名是2015-12-02.jpg这种格式，屏幕显示的日期即为2015.12。
- 支持背景循环播放音频文件。
- 提供降低分辨率载入图像的选项以改善性能。
- 提供单张幻灯片设置（编辑模式），可添加自定义文本。
- 可关联Windows资源管理器右键菜单，右键点击文件夹直接播放。
- 支持EXIF旋转。
- 支持简体中文语言。

## Note
- Press **ESC** to fade out and quit.
- Press **F1** to re-open control panel after it is closed.
- Press **F12** at runtime for options.
- Press **F11** at runtime for Edit Mode dialog.
- Press **Ctrl+P** at runtime to hold off transition to the next image (i.e. pause).
- Press **Shift+P** at runtime to fade out and pause music.
- Press **Ctrl+R** at runtime to restart slideshow.
- Press **Ctrl+Q** to quit immediately without animations (otherwise fade to a black screen).
- Press **Ctrl+N** to jump to the next music item.
- ~~The config.xml file serves as a configuration that is loaded at program start. Paths and other settings can be changed to local / relative path according to the location of the folders on your system.~~
- ~~Please enclose the folder path with `<![CDATA[     ]]>` when the path contains markup chars in XML like &.~~
  - ~~Example: `<![CDATA[F:\Folder with & in name\images]]>`~~
- Animation might not be as smooth when not using Windows Aero themes on Windows 7.
- Preferably a Windows 7 or above PC with a modern discrete GPU gives better / smoother performance. Any configuration that affect Aero performance will also affect animation playback. If you have multiple monitors, set to duplicate or single monitor mode for better performance.
- Framedrops may occur when certain programs are opened such as Potplayer and Foobar 2000.
- Choose "All at Once" under Load Mode to load all images at program start. It uses more memory but eliminates frame-drops in transition animation.
- Audio files with URI escape marks in the file name (e.g. This%20is%20a%20song%28I%20am%20kidding%29.mp3) will not be recognized due to .Net won't take strings in Media.MediaPlayer.Open and the only URI it takes just keeps unescaping the file name whenever it can.

## 使用说明
- 按 **ESC** 键淡出并退出。
- 按 **F1** 键打开/关闭控制窗口。
- 按 **F12** 键打开选项窗口。
- 按 **F11** 键打开编辑模式窗口。
- 按 **Ctrl+P** 可暂停/播放动画。
- 按 **Shift+P** 可淡出并暂停音乐，再按即恢复播放。
- 按 **Ctrl+R** 重新开始幻灯片。
- 按 **Ctrl+Q** 直接跳过动画淡出到桌面（否则则淡出到黑屏）。
- 按 **Ctrl+N** 下一首音乐。
- 在Windows 7系统里，如果Aero特效未开启，动画效果可能受影响。
- 使用独立显卡和Windows 7（或以上）动画效果会更流畅。任何降低Aero性能的配置都会影响动画效果。如果你连接了多个显示器，设置为双显示复制或者只使用一个显示器会改善动画效果。
- 部分应用程序同样会影响帧率。比如Potplayer和Foobar 2000。
- 载入模式选项中，选择“一次性载入”会在程序启动时载入所有图片。此选项会占用大量内存，但会改善换页动画的帧率。
- 文件名中包含URI转义符的音频文件（如：This%20is%20a%20song%28I%20am%20kidding%29.mp3）不会被识别。

## Updates
- 2017-04-23
  - Add bulk details edit in Edit Mode by (selecting multiple images first).
  - Fix initialization logic error.
  - Update author info to compiled executable.
  - Add minimum configuration file version verification.
- 2017-01-14
  - Added fading to desktop animation.
  - Added option to jump to the next music item.
  - Other minor code changes.
- 2017-01-10
  - Added option to avoid showing control window each time the slideshow starts.
  - Added fade-in to the exit prompt.
- 2017-01-07
  - Added option to exclude sub-folders.
  - Minor changes to Options window to fit localizations.
- 2016-10-07
  - Added more text customizations.
- 2016-08-17
  - Added support for reading orientation information from EXIF and rotate image accordingly. Only rotation support is added. EXIF-flipped images (orientation value 2, 4, 5, 7) are less common.
  - Hold returning to desktop during exit by requiring pressing ESC again. (Like in PowerPoint.)
- 2016-08-13
  - Added option to randomly show images.
  - Added option to randomly play audio.
  - Changed some tip texts.
  - Fixed displayed text / date may disappear in the second loop.
- 2016-03-11
  - Added custom text display.
  - Added version verification.
  - Added output preview in Edit Slide window.
  - Improved text / date display loop.
  - Option to add a shortcut menu entry for folders, from where the show can be launched directly.
  - Added fade out while application is closing.
- 2016-03-09
  - Fixed Random transition not working properly.
  - Minor changes on how date animation is handled.
- 2016-03-04
  - Added control window from where pause, restart and other functions can be accessed.
  - Added simplified Chinese language.
  - Minor changes.
- 2016-03-01
  - Added Load Mode option for better performance. Improved transition framerate under All-at-Once mode.
- 2016-02-26
  - **New transition animation "Throw" added!**
  - Changed how certain animations are handled.
  - Added a "Random" transition which randomly use other existing
transitions.
  - Changes in UI and functionality.
- 2016-02-19
  - Added pause and restart control.
  - Added Edit Mode where options for each individual slide can be edited.
  - Added title slide option in Edit Mode.
- 2016-01-20
  - Added more custom options.
- 2016-01-18
  - Fixed breathing not being animated correctly when Vertical Lock is off.
- 2016-01-16
  - **Brand-new transition animation "Breath" added!**
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
