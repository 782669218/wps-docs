**PictureFormat**



包含应用于图片和 OLE 对象的属性和方法。

**说明**

**LinkFormat** 对象只包含应用于链接 OLE 对象的属性和方法。**OLEFormat** 对象包含应用于 OLE 对象的属性和方法，无论这些对象是不是链接对象。

**方法**

|                                                              | 名称                    | 说明                                                         |
| ------------------------------------------------------------ | ----------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IncrementBrightness** | 按指定数值改变图片亮度。可以使用 **Brightness** 属性设置图片的绝对亮度。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IncrementContrast**   | 按指定值改变图片的对比度。可以使用 **Contrast** 属性设置图片的绝对对比度。 |

**属性**

|                                                              | 名称                      | 说明                                                         |
| ------------------------------------------------------------ | ------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**           | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Brightness**            | 返回或设置指定图片或 OLE 对象的亮度。该属性的值必须是 0.0（最暗）到 1.0（最亮）之间的数。**Single** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ColorType**             | 返回或设置应用于指定图片或 OLE 对象的颜色转换类型。**MsoPictureColorType** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Contrast**              | 返回或设置指定图片或 OLE 对象的对比度。该属性的值必须是 0.0（最低对比度）到 1.0（最高对比度）的数。**Single** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**               | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Crop**                  | 返回一个 **Crop** 对象，该对象代表指定的 **PictureFormat** 对象的裁剪设置。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CropBottom**            | 返回或设置从指定图片或 OLE 对象的底部所裁剪下的磅数。**Single** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CropLeft**              | 返回或设置从指定图片或 OLE 对象的左边所裁剪下的磅数。**Single** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CropRight**             | 返回或设置从指定图片或 OLE 对象的右边所裁剪下的磅数。**Single** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CropTop**               | 返回或设置从指定图片或 OLE 对象的顶部所裁剪下的磅数。**Single** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TransparencyColor**     | 返回或设置指定图片的透明色，以红-绿-蓝 (RGB) 值表示。必须将 **TransparentBackground** 属性设置为 **True**，该属性才有效。仅应用于位图。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TransparentBackground** | 使用 **TransparencyColor** 属性可设置透明颜色。仅应用于位图。MsoTriState 类型，可读写。 |

**成员方法**

#### **PictureFormat.IncrementBrightness**

按指定数值改变图片亮度。可以使用 **Brightness** 属性设置图片的绝对亮度。

**语法**

**express.IncrementBrightness(Increment)**

*express*   一个代表 **PictureFormat** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 指定图片的 Brightness 属性值的改变量。正值表示图片变亮，负值表示图片变暗。 |

**说明**

调整图片亮度不能超出 **Brightness** 属性的界限。例如，如果 **Brightness** 属性初始设为 0.9 并将 *Increment* 参数指定为 0.3，则结果亮度为 1.0（**Brightness** 属性的上限）而不是 1.2。

#### **PictureFormat.IncrementContrast**

按指定值改变图片的对比度。可以使用 **Contrast** 属性设置图片的绝对对比度。

**语法**

**express.IncrementContrast(Increment)**

*express*   一个代表 **PictureFormat** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 指定图片的 Contrast 属性值的改变量。正值表示对比度增大，负值表示对比度减小。 |

**说明**

调整图片的对比度不能超出 **Contrast** 属性的上下限。例如，如果 **Contrast** 属性初始设为 0.9 并将 *Increment* 参数指定为 0.3，则结果对比度为 1.0（**Contrast** 属性的上限）而不是 1.2。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例增大 myDocument 上所有对比度未到最大值的图片的对比度。*/ function test(){ 　　　　let shape = Worksheets.Item(1).Shapes        for(let i =1; i<=shape.Count; i++) { 　　　　    if(shape.Item(i).Type == msoPicture || shape.Item(i).Type == msoLinkedPicture) {   　　　　      shape.Item(i).PictureFormat.IncrementContrast(0.1) 　　　　    } 　　　　} }` |

**成员属性**

#### **PictureFormat.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **PictureFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test(){ 　　　　let myObject = ActiveWorkbook        if(myObject.Application.Value == "ET") {     　　　　MsgBox("This is an ET Application object.") 　　　　} 　　　　else {    　　　　 MsgBox("This is not an ET Application object.") 　　　　} }` |

#### **PictureFormat.Brightness**

返回或设置指定图片或 OLE 对象的亮度。该属性的值必须是 0.0（最暗）到 1.0（最亮）之间的数。**Single** 型，可读写。

**语法**

**express.Brightness**

*express*   一个代表 **PictureFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例设置了 myDocument 中第一个形状的亮度。第一个形状必须是图片或 OLE 对象。*/ function test(){ let myDocument = Worksheets.Item(1) myDocument.Shapes.Item(1).PictureFormat.Brightness = 0.3 }` |

#### **PictureFormat.ColorType**

返回或设置应用于指定图片或 OLE 对象的颜色转换类型。**MsoPictureColorType** 类型，可读写。

**语法**

**express.ColorType**

*express*   一个代表 **PictureFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例通过 MsoPictureColorType 将 myDocument 上的第一个形状的颜色变换类型设置为灰度。第一个形状必须是图片或 OLE 对象。*/ function test(){ 　　　　let myDocument = Worksheets.Item(1) 　　　　myDocument.Shapes.Item(1).PictureFormat.ColorType = msoPictureGrayscale }` |

#### **PictureFormat.Contrast**

返回或设置指定图片或 OLE 对象的对比度。该属性的值必须是 0.0（最低对比度）到 1.0（最高对比度）的数。**Single** 型，可读写。

**语法**

**express.Contrast**

*express*   一个代表 **PictureFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例设置了 myDocument 中第一个形状的对比度。第一个形状必须是图片或 OLE 对象。*/ function test(){ let myDocument = Worksheets.Item(1) myDocument.Shapes.Item(1).PictureFormat.Contrast = 0.8 }` |

#### **PictureFormat.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **PictureFormat** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **PictureFormat.Crop**

返回一个 **Crop** 对象，该对象代表指定的 **PictureFormat** 对象的裁剪设置。

**语法**

**express.Crop**

*express*   一个代表 **PictureFormat** 对象的变量。

#### **PictureFormat.CropBottom**

返回或设置从指定图片或 OLE 对象的底部所裁剪下的磅数。**Single** 型，可读写。

**语法**

**express.CropBottom**

*express*   一个代表 **PictureFormat** 对象的变量。

**说明**

裁剪的计算是同图像的原始尺寸相关的。例如，如果插入一幅原始高度为 100 磅的图像，将其放大为 200 磅高，然后设置 **CropBottom** 属性为 50, 那么从图像的底部切除的将是 100 磅（而不是 50 磅）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例在 myDocument 中第三个形状的底部裁剪了 20 磅。要使此示例执行，第三个形状必须是图片或 OLE 对象。*/ function test(){ 　　　　let myDocument = Worksheets.Item(1) 　　　　myDocument.Shapes.Item(3).PictureFormat.CropBottom = 20 }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*使用此示例可指定从选定的形状底部裁去的百分比，不管是否已调整了该形状的大小。为使此示例正常运行，选定的形状必须是图片或 OLE 对象。*/ function test(){ 　　　　let percentToCrop = prompt( "What percentage do you want to crop off " + " the bottom of this picture?") 　　　　let shapeToCrop = ActiveWindow.Selection.ShapeRange.Item(1) 　　　　let du = shapeToCrop.Duplicate() 　　　　du.ScaleHeight(1, true,du.Height) 　　　　du.Delete() 　　　　let cropPoints = origHeight * percentToCrop / 100 　　　　shapeToCrop.PictureFormat.CropBottom = cropPoints }` |

#### **PictureFormat.CropLeft**

返回或设置从指定图片或 OLE 对象的左边所裁剪下的磅数。**Single** 型，可读写。

**语法**

**express.CropLeft**

*express*   一个代表 **PictureFormat** 对象的变量。

**说明**

裁剪的计算是同图像的原始尺寸相关的。例如，如果插入一幅原始宽度为 100 磅的图像，将其放大为 200 磅宽，然后设置 **CropLeft** 属性为 50, 那么 100 磅（不是 50 磅）将从图像的左边被切除。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例在 myDocument 中第三个形状的左侧裁剪了 20 磅。要使此示例执行，第三个形状必须是图片或 OLE 对象。*/ function test(){ 　　　　let myDocument = Worksheets.Item(1) 　　　　myDocument.Shapes.Item(3).PictureFormat.CropLeft = 20 }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*使用此示例可指定从选定的形状左边裁去的百分比，不管是否已调整了该形状的大小。为使此示例正常运行，选定的形状必须是图片或 OLE 对象。*/ function test(){ 　　　　let percentToCrop = prompt( "What percentage do you want to crop off " + " the left of this picture?") 　　　　let shapeToCrop = ActiveWindow.Selection.ShapeRange.Item(1) 　　　　let du = shapeToCrop.Duplicate() 　　　　du.ScaleHeight(1, true,du.Width) 　　　　du.Delete() 　　　　let cropPoints = origHeight * percentToCrop / 100 　　　　shapeToCrop.PictureFormat.CropBottom = cropPoints }` |

#### **PictureFormat.CropRight**

返回或设置从指定图片或 OLE 对象的右边所裁剪下的磅数。**Single** 型，可读写。

**语法**

**express.CropRight**

*express*   一个代表 **PictureFormat** 对象的变量。

**说明**

裁剪的计算是同图像的原始尺寸相关的。例如，如果插入一幅原始宽度为 100 磅的图像，将其放大为 200 磅宽，然后设置 **CropRight** 属性为 50, 那么 100 磅（不是 50 磅）将从图像的右边被切除。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例在 myDocument 中第三个形状的右侧裁剪了 20 磅。要使此示例执行，第三个形状必须是图片或 OLE 对象。*/ function test(){ 　　　　let myDocument = Worksheets.Item(1) 　　　　myDocument.Shapes.Item(3).PictureFormat.CropRight = 20 }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*使用此示例可指定从选定的形状右边裁去的百分比，不管是否已调整了该形状的大小。为使此示例正常运行，选定的形状必须是图片或 OLE 对象。*/ function test(){ 　　　　let percentToCrop = prompt( "What percentage do you want to crop  " + " off the right of this picture?") 　　　　let shapeToCrop = ActiveWindow.Selection.ShapeRange.Item(1) 　　　　let du = shapeToCrop.Duplicate() 　　　　du.ScaleHeight(1, true,du.Width) 　　　　du.Delete() 　　　　let cropPoints = origHeight * percentToCrop / 100 　　　　shapeToCrop.PictureFormat.CropBottom = cropPoints }` |

#### **PictureFormat.CropTop**

返回或设置从指定图片或 OLE 对象的顶部所裁剪下的磅数。**Single** 型，可读写。

**语法**

**express.CropTop**

*express*   一个代表 **PictureFormat** 对象的变量。

**说明**

裁剪的计算是同图像的原始尺寸相关的。例如，如果插入一幅原始高度为 100 磅的图像，将其放大为 200 磅高，然后设置 **CropTop** 属性为 50, 那么 100 磅（不是 50 磅）将从图像的顶部被切除。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例在 myDocument 中第三个形状的顶部裁剪了 20 磅。要使此示例执行，第三个形状必须是图片或 OLE 对象。*/ function test(){ 　　　　let myDocument = Worksheets.Item(1) 　　　　myDocument.Shapes.Item(3).PictureFormat.CropTop = 20 }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*使用此示例可指定从选定的形状顶部裁去的百分比，不管是否已调整了该形状的大小。为使此示例正常运行，选定的形状必须是图片或 OLE 对象。*/ function test(){ 　　　　let percentToCrop = prompt( "What percentage do you want to crop " + "  off the top of this picture?") 　　　　let shapeToCrop = ActiveWindow.Selection.ShapeRange.Item(1) 　　　　let du = shapeToCrop.Duplicate() 　　　　du.ScaleHeight(1, true,du.Height) 　　　　du.Delete() 　　　　let cropPoints = origHeight * percentToCrop / 100 　　　　shapeToCrop.PictureFormat.CropBottom = cropPoints }` |

#### **PictureFormat.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **PictureFormat** 对象的变量。

#### **PictureFormat.TransparencyColor**

返回或设置指定图片的透明色，以红-绿-蓝 (RGB) 值表示。必须将 **TransparentBackground** 属性设置为 **True**，该属性才有效。仅应用于位图。**Long** 类型，可读写。

**语法**

**express.TransparencyColor**

*express*   一个代表 **PictureFormat** 对象的变量。

**说明**

如果想透过图片的透明部分始终看到图片后面的对象，必须将图片的 **FillFormat** 对象的 **Visible** 属性设为 **False**。如果图片有透明颜色且 **FillFormat** 对象的 **Visible** 属性被设为 **True**，则图片的填充可以透过透明颜色看到，但图片后面的对象将难以看到。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本例将 myDocument 的第一个形状中带有 RGB 值的颜色设置为透明色，该颜色由 RGB(0, 0, 255) 函数返回。要使本示例执行，第一个形状必须是位图。*/ function test(){ 　　　　let blueScreen = (0,0,255) 　　　　let myDocument = Worksheets.Item(1) 　　　　let myshape = myDocument.Shapes.Item(1)   　　　let mypicture = myshape.PictureFormat      　　　　   mypicture.TransparentBackground = true      　　　　   mypicture.TransparencyColor = blueScreen  　　　　   myshape.Fill.Visible = false }` |

#### **PictureFormat.TransparentBackground**

使用 **TransparencyColor** 属性可设置透明颜色。仅应用于位图。MsoTriState 类型，可读写。

**语法**

**express.TransparentBackground**

*express*   一个代表 **PictureFormat** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。    |
| ------------------------------------------------ |
| msoCTrue                                         |
| msoFalse                                         |
| msoTriStateMixed                                 |
| msoTriStateToggle                                |
| msoTrue 图片中定义为透明颜色的部分显示为透明的。 |

如果想透过图片的透明部分始终看到图片后面的对象，必须将图片的 **FillFormat** 对象的 **Visible** 属性设为 **False**。如果图片有透明颜色且 **FillFormat** 对象的 **Visible** 属性被设为 **True**，则图片的填充可以透过透明颜色看到，但图片后面的对象将难以看到。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本例将 myDocument 的第一个形状中带有 RGB 值的颜色设置为透明色，该颜色由 RGB(0, 24, 240) 函数返回。要使本示例执行，第一个形状必须是位图。 */ function test(){ 　　　　let blueScreen = (0,0,255) 　　　　let myDocument = Worksheets.Item(1) 　　　　let myshape = myDocument.Shapes.Item(1)   　　　let mypicture = myshape.PictureFormat       　　　　　mypicture.TransparentBackground = true         　　　　mypicture.TransparencyColor = blueScreen     　　　　myshape.Fill.Visible = false }` |

适用环境：web

适用平台：windows/linux