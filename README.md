# eVB-ImageListEx
An eVB implementation of the ImageList control, with support for masked (transparent) images.

This is a module and accompanying test project to directly work with Windows image lists. This module can fully replace the eVB ImageList control.

The most notable improvement is support for masked images, so it's possible to use icons with transparent backgrounds on the [taskbar](https://github.com/WinCEDev/eVB-Taskbar-Icon) and [notification area](https://github.com/WinCEDev/eVB-Notification-Icon), as well as any of the native or third-party eVB controls that take a handle to an image list.

However even if you don't need transparency you can potentially save around 50kb and any COM registration headaches when distributing your app by not having to include the MSCEImageList.dll control. This is especially useful if you only intend to use it for a taskbar or notification area icon.

## Usage

A complete application could look like this:

```vb
'A CommandBar control has been added to the form.

Private CommandBarIcons As Long 'This holds the image list handle.

Private Sub Form_Load()

    'If you need support for transparency, set flags to 'ILC_COLOR Or ILC_MASK', otherwise, you can set it to just 'ILC_COLOR'.
    CommandBarIcons = ImageListEx_Create(16, 16, ILC_COLOR Or ILC_MASK) 'Create the image list.

    'Add the toolbar bitmap to the image list, we use magenta as our transparent color.
    ImageListEx_AddMasked CommandBarIcons, "toolbar.bmp", vbMagenta

    'Assign our image list handle to the CommandBar.
    CommandBar.ImageList = CommandBarIcons

    'At this point our image list is loaded and we can add a button as per normal.
    Dim objButton As CommandBarButton

    Set objButton = CommandBar.Controls.Add(cbrButton)
    objButton.Image = 0 'Assign our image.

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Make sure to always destroy the image list when the form containing it closes or the application ends.
    ImageListEx_Destroy CommandBarIcons
End Sub
```

## Screenshots

![Screenshot showing the example application in the Maple color scheme.](https://github.com/WinCEDev/eVB-ImageListEx/blob/main/Screenshots/CAPT0000.png?raw=1)

![Screenshot showing the example application in the Spruce color scheme.](https://github.com/WinCEDev/eVB-ImageListEx/blob/main/Screenshots/CAPT0001.png?raw=1)

## Links

- [HPC:Factor Forum Thread](https://www.hpcfactor.com/forums/forums/thread-view.asp?tid=20876&posts=1)
