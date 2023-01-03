**PictureFormat**



包含应用于图片和 OLE 对象的属性和方法。

**说明**

**LinkFormat** 对象包含仅应用于链接 OLE 对象的属性和方法。**OLEFormat** 对象包含应用于链接和非链接 OLE 对象的属性和方法。

**方法**

|                                                              | 名称                    | 说明                                                         |
| ------------------------------------------------------------ | ----------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IncrementBrightness** | 按指定数值更改图片的亮度。可以使用 **Brightness** 属性设置图片的绝对亮度。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IncrementContrast**   | 按指定数值更改图片的对比度。可以使用 **Contrast** 属性设置图片的绝对对比度。 |

**属性**

|                                                              | 名称                      | 说明                                                         |
| ------------------------------------------------------------ | ------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**           | 返回一个 **Application** 对象，该对象表示指定对象的创建者。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Brightness**            | 返回或设置指定图片或 OLE 对象的亮度。该属性的值必须是 0.0（最暗）到 1.0（最亮）之间的数。可读/写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ColorType**             | 返回或设置应用于指定图片或 OLE 对象的颜色转换类型。可读/写 **MsoPictureColorType** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Contrast**              | 返回或设置指定图片或 OLE 对象的对比度。该属性的值必须是 0.0（最低对比度）到 1.0（最高对比度）的数。可读/写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**               | 返回 **Long** 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CropBottom**            | 返回或设置从指定图片或 OLE 对象的底部裁剪下的磅值。可读/写 **Single** 类型。裁剪以相对于图片的原始尺寸计算。例如，如果插入的图片原始高度为 100 磅，将其高度调整为 200 磅，再将 **CropBottom** 属性设置为 50，则会从图片底部裁剪 100 磅（而不是 50 镑）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CropLeft**              | 返回或设置从指定图片或 OLE 对象的左侧裁剪下的磅值。可读/写 **Single** 类型。裁剪以相对于图片的原始尺寸计算。例如，如果插入的图片原始宽度为 100 磅，将其宽度调整为 200 磅，再将 **CropLeft** 属性设置为 50，则会从图片左侧裁剪 100 磅（而不是 50 磅）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CropRight**             | 返回或设置从指定图片或 OLE 对象的右侧裁剪下的磅值。可读/写 **Single** 类型。裁剪以相对于图片的原始尺寸计算。例如，如果插入的图片原始宽度为 100 磅，将其宽度调整为 200 磅，再将 **CropRight** 属性设置为 50，则会从图片右侧裁剪 100 磅（而不是 50 磅）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CropTop**               | 返回或设置从指定图片或 OLE 对象的顶部裁剪下的磅值。可读/写 **Single** 类型。裁剪以相对于图片的原始尺寸计算。例如，如果插入的图片原始高度为 100 磅，将其高度调整为 200 磅，再将 **CropTop** 属性设置为 50，则会从图片顶部裁剪 100 磅（而不是 50 镑）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                | 返回指定对象的父对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TransparencyColor**     | 以红-绿-蓝 (RGB) 值返回或设置指定图片的透明色。要使该属性生效，必须将 **TransparentBackground** 属性设为 **True**。仅适用于位图。可读/写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TransparentBackground** | 决定图片中定义为透明色的部分是否变为透明。可读/写 **MsoTriState** 类型。仅适用于位图。 |

**成员方法**

#### **PictureFormat.IncrementBrightness**

按指定数值更改图片的亮度。可以使用 **Brightness** 属性设置图片的绝对亮度。

**语法**

**express.IncrementBrightness(Increment)**

*express*   一个代表 **PictureFormat** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 指定图片 Brightness 属性值的改变量。正值增加图片亮度；负值降低图片亮度。 |

**说明**

图片的亮度调整不能超出 **Brightness** 属性的上下限。例如，如果将 **Brightness** 属性初始设为 0.9，并指定 **Increment** 参数为 0.3，则最后的亮度不是 1.2，而是 **Brightness** 属性的上限，即 1.0。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建 myDocument 上第一个形状的副本，再移动该副本并使之变暗。要执行本示例，第一个形状必须是图片或 OLE 对象。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let mypane = myDocument.Shapes.Item(1).Duplicate()  　　　　   mypane.PictureFormat.IncrementBrightness(-0.2)   　　　　  mypane.IncrementLeft(50)  　　　　   mypane.IncrementTop(50) }` |

#### **PictureFormat.IncrementContrast**

按指定数值更改图片的对比度。可以使用 **Contrast** 属性设置图片的绝对对比度。

**语法**

**express.IncrementContrast(Increment)**

*express*   一个代表 **PictureFormat** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 指定图片的 Contrast 属性值的改变量。正值表示对比度增大，负值表示对比度减小。 |

**说明**

图片的对比度调整不能超出 **Contrast** 属性的上下限。例如，如果将 **Contrast** 属性初始设为 0.9，并指定 **Increment** 参数为 0.3，则最后的对比度不是 1.2，而是 **Contrast** 属性的上限，即 1.0。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对 myDocument 上还没有设置为最大对比度的所有图片增加对比度。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　for(let i = 1; i <= myDocument.Shapes.Count; i++){    　　　　 if(myDocument.Shapes.Item(i).Type == msoPicture){      　　　　  myDocument.Shapes.Item(i).PictureFormat.IncrementContrast(0.1)  　　　　   } 　　　　} }` |

**成员属性**

#### **PictureFormat.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **PictureFormat** 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function AddAndSave(pptPres){     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes     for (let i = 1; i <= shpOle.Count; i++) {         if (shpOle.Item(i).Type == msoLinkedOLEObject) {             alert(shpOle.Item(i).OLEFormat.Application.Name)         }     } }` |

#### **PictureFormat.Brightness**

返回或设置指定图片或 OLE 对象的亮度。该属性的值必须是 0.0（最暗）到 1.0（最亮）之间的数。可读/写 **Single** 类型。

**语法**

**express.Brightness**

*express*   一个代表 **PictureFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置了 myDocument 中第一个形状的亮度。第一个形状必须为图片或 OLE 对象。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)    　　　　 myDocument.Shapes.Item(1).PictureFormat.Brightness = 0.3 }` |

#### **PictureFormat.ColorType**

返回或设置应用于指定图片或 OLE 对象的颜色转换类型。可读/写 **MsoPictureColorType** 类型。

**语法**

**express.ColorType**

*express*   一个代表 **PictureFormat** 对象的变量。

**说明**

| MsoPictureColorType 可以是下列 MsoPictureColorType 常量之一。 |
| ------------------------------------------------------------ |
| msoPictureAutomatic                                          |
| msoPictureBlackAndWhite                                      |
| msoPictureGrayscale                                          |
| msoPictureMixed                                              |
| msoPictureWatermark                                          |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 myDocument 上的第一个形状的颜色变换方式设置为灰度级。第一个形状必须为图片或 OLE 对象。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　myDocument.Shapes.Item(1).PictureFormat.ColorType = msoPictureGrayscale }` |

#### **PictureFormat.Contrast**

返回或设置指定图片或 OLE 对象的对比度。该属性的值必须是 0.0（最低对比度）到 1.0（最高对比度）的数。可读/写 **Single** 类型。

**语法**

**express.Contrast**

*express*   一个代表 **PictureFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置了 myDocument 中第一个形状的对比度。第一个形状必须为图片或 OLE 对象。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　myDocument.Shapes.Item(1).PictureFormat.Contrast = 0.8 }` |

#### **PictureFormat.Creator**

返回 **Long** 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。

**语法**

**express.Creator**

*express*   一个代表 **PictureFormat** 对象的变量。

**说明**

**Creator** 属性用于 Macintosh 下的 WPS Office应用程序。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示关于 myObject 创建者的消息。*/ function test(){ 　　　　let myObject = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1)  　　　　if(myObject.Creator == parseInt(50575054, 16)){ 　　　　   alert("This is aWPP object") 　　　　} 　　　　else{ 　　　　   alert("This is not aWPP object") 　　　　} }` |

#### **PictureFormat.CropBottom**

返回或设置从指定图片或 OLE 对象的底部裁剪下的磅值。可读/写 **Single** 类型。裁剪以相对于图片的原始尺寸计算。例如，如果插入的图片原始高度为 100 磅，将其高度调整为 200 磅，再将 **CropBottom** 属性设置为 50，则会从图片底部裁剪 100 磅（而不是 50 镑）。

**语法**

**express.CropBottom**

*express*   一个代表 **PictureFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 myDocument 中第三个形状的底部裁剪了 20 磅。要使本示例执行，第三个形状必须是图片或 OLE 对象。*/ function test(){ 　　　　let myDocument = ActivePresentation.Slides.Item(1) 　　　　myDocument.Shapes.Item(3).PictureFormat.CropBottom = 20 }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将选定形状的底部裁剪用户指定的百分比，而不考虑该形状是否调整过。要使本示例执行，选定的形状必须是图片或 OLE 对象。*/ function test(){ 　　　　let percentToCrop = Application.prompt("What percentage do you want to crop off the bottom of this picture?") 　　　　let shapeToCrop = Application.ActiveWindow.Selection.ShapeRange.Item(1)  　　　　let myPictureFormat = shapeToCrop.Duplicate() 　　　　    myPictureFormat.ScaleHeight(1, true)  　　　　let origHeight = myPictureFormat.Height 　　　　myPictureFormat.Delete()  　　　　let cropPoints = origHeight * percentToCrop / 100 　　　　shapeToCrop.PictureFormat.CropBottom = cropPoints }` |

#### **PictureFormat.CropLeft**

返回或设置从指定图片或 OLE 对象的左侧裁剪下的磅值。可读/写 **Single** 类型。裁剪以相对于图片的原始尺寸计算。例如，如果插入的图片原始宽度为 100 磅，将其宽度调整为 200 磅，再将 **CropLeft** 属性设置为 50，则会从图片左侧裁剪 100 磅（而不是 50 磅）。

**语法**

**express.CropLeft**

*express*   一个代表 **PictureFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 myDocument 中第三个形状的左侧裁剪了 20 磅。要使本示例执行，第三个形状必须是图片或 OLE 对象。*/ function test(){ 　　　　let myDocument = ActivePresentation.Slides.Item(1) 　　　　myDocument.Shapes.Item(3).PictureFormat.CropLeft = 20 }` |

#### **PictureFormat.CropRight**

返回或设置从指定图片或 OLE 对象的右侧裁剪下的磅值。可读/写 **Single** 类型。裁剪以相对于图片的原始尺寸计算。例如，如果插入的图片原始宽度为 100 磅，将其宽度调整为 200 磅，再将 **CropRight** 属性设置为 50，则会从图片右侧裁剪 100 磅（而不是 50 磅）。

**语法**

**express.CropRight**

*express*   一个代表 **PictureFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 myDocument 中第三个形状的右侧裁剪了 20 磅。要使本示例执行，第三个形状必须是图片或 OLE 对象。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　myDocument.Shapes.Item(3).PictureFormat.CropRight = 20  }` |

#### **PictureFormat.CropTop**

返回或设置从指定图片或 OLE 对象的顶部裁剪下的磅值。可读/写 **Single** 类型。裁剪以相对于图片的原始尺寸计算。例如，如果插入的图片原始高度为 100 磅，将其高度调整为 200 磅，再将 **CropTop** 属性设置为 50，则会从图片顶部裁剪 100 磅（而不是 50 镑）。

**语法**

**express.CropTop**

*express*   一个代表 **PictureFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 myDocument 中第三个形状的顶部裁剪了 20 磅。要使本示例执行，第三个形状必须是图片或 OLE 对象。*/ function test(){ let myDocument = Application.ActivePresentation.Slides.Item(1) myDocument.Shapes.Item(3).PictureFormat.CropTop = 20 }` |

#### **PictureFormat.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **PictureFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test(){ 　　　　let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes 　　　　let myPicture = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame    　　　　 myPicture.TextRange.Text = "Test text"     　　　　myPicture.Parent.Rotation = 45 }` |

#### **PictureFormat.TransparencyColor**

以红-绿-蓝 (RGB) 值返回或设置指定图片的透明色。要使该属性生效，必须将 **TransparentBackground** 属性设为 **True**。仅适用于位图。可读/写 **Long** 类型。

**语法**

**express.TransparencyColor**

*express*   一个代表 **PictureFormat** 对象的变量。

**说明**

如果想透过图片的透明部分始终看到图片后面的对象，必须将图片的 **FillFormat** 对象的 **Visible** 属性设为 **False**。如果图片有透明色且其 **FillFormat** 对象的 **Visible** 属性被设为 **True**，则可以透过透明色看到图片的填充，但图片后面的对象将难以看到。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例用函数 RGB(0, 0, 255) 返回 RGB 值，并将该 RGB 值对应的颜色设置为 myDocument 上第一个形状的透明色。要使本示例正常运行，第一个形状必须是位图。*/ function test(){ 　　　　let blueScreen = (0, 0, 255) 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let myPicture = myDocument.Shapes.Item(1).PictureFormat  　　　　   myPicture.TransparentBackground = true  　　　　   myPicture.TransparencyColor = blueScreen 　　　　myDocument.Shapes.Item(1).Fill.Visible = false }` |

#### **PictureFormat.TransparentBackground**

决定图片中定义为透明色的部分是否变为透明。可读/写 **MsoTriState** 类型。仅适用于位图。

**语法**

**express.TransparentBackground**

*express*   一个代表 **PictureFormat** 对象的变量。

**说明**

使用 **TransparencyColor** 属性可设置透明色。

如果想透过图片的透明部分始终看到图片后面的对象，必须将图片的 **FillFormat** 对象的 **Visible** 属性设为 **msoFalse**。如果图片有透明色且其 **FillFormat** 对象的 **Visible** 属性被设为 **msoTrue**，则图片的填充可以透过透明色看到，但图片后面的对象将难以看到。

| MsoTriState 可以是下列 MsoTriState 常量之一。  |
| ---------------------------------------------- |
| msoCTrue                                       |
| msoFalse                                       |
| msoTriStateMixed                               |
| msoTriStateToggle                              |
| msoTrue 图片中颜色定义为透明色的部分变为透明。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例用函数 RGB(0, 24, 240) 返回 RGB 值，并将该 RGB 值对应的颜色设置为 myDocument 上第一个形状的透明色。要使本示例正常运行，第一个形状必须是位图。*/ function test(){ 　　　　let blueScreen = (0, 0, 255) 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let myPicture = myDocument.Shapes.Item(1).PictureFormat  　　　　   myPicture.TransparentBackground = true  　　　　   myPicture.TransparencyColor = blueScreen 　　　　myDocument.Shapes.Item(1).Fill.Visible = false }` |

适用环境：web

适用平台：windows/linux