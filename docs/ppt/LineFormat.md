**LineFormat**



该对象代表线条和箭头格式。对于线条，**LineFormat** 对象包含关于线条本身的格式信息；对于带有边框的图形，本对象包含该图形的边框的格式信息。

**属性**

|                                                              | 名称                     | 说明                                                         |
| ------------------------------------------------------------ | ------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**          | 返回一个 **Application** 对象，该对象表示指定对象的创建者。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BackColor**            | 返回或设置 **ColorFormat** 对象，该对象代表指定填充对象或图案线条的背景色。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BeginArrowheadLength** | 返回或设置指定线条始端的箭头长度。可读/写 **MsoArrowheadLength** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BeginArrowheadStyle**  | 返回或设置指定线条始端的箭头样式。可读/写 **MsoArrowheadStyle** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BeginArrowheadWidth**  | 返回或设置指定线条始端的箭头宽度。可读/写 **MsoArrowheadWidth** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**              | 返回 **Long** 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DashStyle**            | 返回或设置指定线条的虚线线型。可读/写 **MsoLineDashStyle** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EndArrowheadLength**   | 返回或设置指定线条末端箭头的长度。可读/写 **MsoArrowheadLength** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EndArrowheadStyle**    | 返回或设置指定线条末端箭头的样式。可读/写 **MsoArrowheadStyle** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EndArrowheadWidth**    | 返回或设置指定线条末端箭头的宽度。可读/写 **MsoArrowheadWidth** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ForeColor**            | 返回或设置 **ColorFormat** 对象，该对象表示填充、线条或阴影的前景色。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InsetPen**             | 该属性值为 **MsoTrue** 时，在指定形状的内部绘制线条。可读/写 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**               | 返回指定对象的父对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Pattern**              | 设置或返回一个值，该值代表应用于指定线条的图案。**MsoPatternType** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Style**                | 返回或设置线条样式。可读/写 **MsoLineStyle** 类型。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Transparency**         | 该属性返回或设置指定填充、阴影或线条的透明度，可取 0.0（不透明）至 1.0（透明）之间的数值。**Single** 类型，可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**              | 返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。**MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Weight**               | 以磅为单位返回或设置指定线条的粗细。可读/写 **Single** 类型。 |

**成员属性**

#### **LineFormat.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **LineFormat** 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function AddAndSave(pptPres) {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes     for (let i = 1; i <= shpOle.Count; i++) {         if (shpOle.Item(i).Type == msoLinkedOLEObject) {             alert(shpOle.Item(i).OLEFormat.Application.Name)         }     } }` |

#### **LineFormat.BackColor**

返回或设置 **ColorFormat** 对象，该对象代表指定填充对象或图案线条的背景色。可读/写。

**语法**

**express.BackColor**

*express*   一个代表 **LineFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例向 myDocument 中添加一个矩形，然后设置该矩形的填充的前景色、背景色和渐变。*/ function test(){ 　　　　let myDocument = ActivePresentation.Slides.Item(1) 　　　　let fill = myDocument.Shapes.AddShape(msoShapeRectangle,90, 90, 90, 50).Fill 　　　　fill.ForeColor.RGB = (128, 0, 0) 　　　　fill.BackColor.RGB = (170, 170, 170) 　　　　fill.TwoColorGradient(msoGradientHorizontal, 1) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将一条带图案的线条添加至 myDocument。*/ function test() {     let myDocument = Application.ActivePresentation.Slides.Item(1)     let line = myDocument.Shapes.AddLine(10, 100, 250, 0).Line     line.Weight = 6     line.ForeColor.RGB = (0, 0, 255)     line.BackColor.RGB = (128, 0, 0)     line.Pattern = msoPatternDarkDownwardDiagonal }` |

#### **LineFormat.BeginArrowheadLength**

返回或设置指定线条始端的箭头长度。可读/写 **MsoArrowheadLength** 类型。

**语法**

**express.BeginArrowheadLength**

*express*   一个代表 **LineFormat** 对象的变量。

**说明**

| MsoArrowheadLength 可以是下列 MsoArrowheadLength 常量之一。 |
| ----------------------------------------------------------- |
| msoArrowheadLengthMedium                                    |
| msoArrowheadLengthMixed                                     |
| msoArrowheadLong                                            |
| msoArrowheadShort                                           |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加一条线。在该线条的起点处有一个短而窄的椭圆，在线条的终点处有一个长而宽的三角形。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line 　　　　line.BeginArrowheadLength = msoArrowheadShort 　　　　line.BeginArrowheadStyle = msoArrowheadOval 　　　　line.BeginArrowheadWidth = msoArrowheadNarrow 　　　　line.EndArrowheadLength = msoArrowheadLong 　　　　line.EndArrowheadStyle = msoArrowheadTriangle 　　　　line.EndArrowheadWidth = msoArrowheadWide }` |

#### **LineFormat.BeginArrowheadStyle**

返回或设置指定线条始端的箭头样式。可读/写 **MsoArrowheadStyle** 类型。

**语法**

**express.BeginArrowheadStyle**

*express*   一个代表 **LineFormat** 对象的变量。

**说明**

| MsoArrowheadStyle 可以是下列 MsoArrowheadStyle 常量之一。 |
| --------------------------------------------------------- |
| msoArrowheadDiamond                                       |
| msoArrowheadNone                                          |
| msoArrowheadOpen                                          |
| msoArrowheadOval                                          |
| msoArrowheadStealth                                       |
| msoArrowheadStyleMixed                                    |
| msoArrowheadTriangle                                      |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加一条线。在该线条的起点处有一个短而窄的椭圆，在线条的终点处有一个长而宽的三角形。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line 　　　　line.BeginArrowheadLength = msoArrowheadShort 　　　　line.BeginArrowheadStyle = msoArrowheadOval 　　　　line.BeginArrowheadWidth = msoArrowheadNarrow 　　　　line.EndArrowheadLength = msoArrowheadLong 　　　　line.EndArrowheadStyle = msoArrowheadTriangle 　　　　line.EndArrowheadWidth = msoArrowheadWide }` |

#### **LineFormat.BeginArrowheadWidth**

返回或设置指定线条始端的箭头宽度。可读/写 **MsoArrowheadWidth** 类型。

**语法**

**express.BeginArrowheadWidth**

*express*   一个代表 **LineFormat** 对象的变量。

**说明**

| MsoArrowheadWidth 可以是下列 MsoArrowheadWidth 常量之一。 |
| --------------------------------------------------------- |
| msoArrowheadNarrow                                        |
| msoArrowheadWide                                          |
| msoArrowheadWidthMedium                                   |
| msoArrowheadWidthMixed                                    |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加一条线。在该线条的起点处有一个短而窄的椭圆，在线条的终点处有一个长而宽的三角形。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line 　　　　line.BeginArrowheadLength = msoArrowheadShort 　　　　line.BeginArrowheadStyle = msoArrowheadOval 　　　　line.BeginArrowheadWidth = msoArrowheadNarrow 　　　　line.EndArrowheadLength = msoArrowheadLong 　　　　line.EndArrowheadStyle = msoArrowheadTriangle 　　　　line.EndArrowheadWidth = msoArrowheadWide }` |

#### **LineFormat.Creator**

返回 **Long** 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。

**语法**

**express.Creator**

*express*   一个代表 **LineFormat** 对象的变量。

**说明**

**Creator** 属性用于 Macintosh 下的 WPS Office应用程序

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示关于 myObject 创建者的消息。*/ function test(){ 　　　　let myObject = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1) 　　　　if (myObject.Creator == "0x50575054") { 　　　　    alert("This is aWPP object") 　　　　} 　　　　else { 　　　　    alert("This is not aWPP object") 　　　　} }` |

#### **LineFormat.DashStyle**

返回或设置指定线条的虚线线型。可读/写 **MsoLineDashStyle** 类型。

**语法**

**express.DashStyle**

*express*   一个代表 **LineFormat** 对象的变量。

**说明**

| MsoLineDashStyle 可以是下列 MsoLineDashStyle 常量之一。 |
| ------------------------------------------------------- |
| msoLineDash                                             |
| msoLineDashDot                                          |
| msoLineDashDotDot                                       |
| msoLineDashStyleMixed                                   |
| msoLineLongDash                                         |
| msoLineLongDashDot                                      |
| msoLineRoundDot                                         |
| msoLineSolid                                            |
| msoLineSquareDot                                        |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例向 myDocument 中添加一条蓝色虚线。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let line = myDocument.Shapes.AddLine(10, 10, 250, 250).Line 　　　　line.DashStyle = msoLineDashDotDot 　　　　line.ForeColor.RGB = (50, 0, 128) }` |

#### **LineFormat.EndArrowheadLength**

返回或设置指定线条末端箭头的长度。可读/写 **MsoArrowheadLength** 类型。

**语法**

**express.EndArrowheadLength**

*express*   一个代表 **LineFormat** 对象的变量。

**说明**

| MsoArrowheadLength 可以是下列 MsoArrowheadLength 常量之一。 |
| ----------------------------------------------------------- |
| msoArrowheadLengthMedium                                    |
| msoArrowheadLengthMixed                                     |
| msoArrowheadLong                                            |
| msoArrowheadShort                                           |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加一条线。在该线条的起点处有一个短而窄的椭圆，在线条的终点处有一个长而宽的三角形。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line 　　　　line.BeginArrowheadLength = msoArrowheadShort 　　　　line.BeginArrowheadStyle = msoArrowheadOval 　　　　line.BeginArrowheadWidth = msoArrowheadNarrow 　　　　line.EndArrowheadLength = msoArrowheadLong 　　　　line.EndArrowheadStyle = msoArrowheadTriangle 　　　　line.EndArrowheadWidth = msoArrowheadWide }` |

#### **LineFormat.EndArrowheadStyle**

返回或设置指定线条末端箭头的样式。可读/写 **MsoArrowheadStyle** 类型。

**语法**

**express.EndArrowheadStyle**

*express*   一个代表 **LineFormat** 对象的变量。

**说明**

| MsoArrowheadStyle 可以是下列 MsoArrowheadStyle 常量之一。 |
| --------------------------------------------------------- |
| msoArrowheadDiamond                                       |
| msoArrowheadNone                                          |
| msoArrowheadOpen                                          |
| msoArrowheadOval                                          |
| msoArrowheadStealth                                       |
| msoArrowheadStyleMixed                                    |
| msoArrowheadTriangle                                      |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加一条线。在该线条的起点处有一个短而窄的椭圆，在线条的终点处有一个长而宽的三角形。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line 　　　　line.BeginArrowheadLength = msoArrowheadShort 　　　　line.BeginArrowheadStyle = msoArrowheadOval 　　　　line.BeginArrowheadWidth = msoArrowheadNarrow 　　　　line.EndArrowheadLength = msoArrowheadLong 　　　　line.EndArrowheadStyle = msoArrowheadTriangle 　　　　line.EndArrowheadWidth = msoArrowheadWide }` |

#### **LineFormat.EndArrowheadWidth**

返回或设置指定线条末端箭头的宽度。可读/写 **MsoArrowheadWidth** 类型。

**语法**

**express.EndArrowheadWidth**

*express*   一个代表 **LineFormat** 对象的变量。

**说明**

| MsoArrowheadWidth 可以是下列 MsoArrowheadWidth 常量之一。 |
| --------------------------------------------------------- |
| msoArrowheadNarrow                                        |
| msoArrowheadWide                                          |
| msoArrowheadWidthMedium                                   |
| msoArrowheadWidthMixed                                    |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加一条线。在该线条的起点处有一个短而窄的椭圆，在线条的终点处有一个长而宽的三角形。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line 　　　　line.BeginArrowheadLength = msoArrowheadShort 　　　　line.BeginArrowheadStyle = msoArrowheadOval 　　　　line.BeginArrowheadWidth = msoArrowheadNarrow 　　　　line.EndArrowheadLength = msoArrowheadLong 　　　　line.EndArrowheadStyle = msoArrowheadTriangle 　　　　line.EndArrowheadWidth = msoArrowheadWide }` |

#### **LineFormat.ForeColor**

返回或设置 **ColorFormat** 对象，该对象表示填充、线条或阴影的前景色。可读/写。

**语法**

**express.ForeColor**

*express*   一个代表 **LineFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例向 myDocument 中添加一个矩形，然后设置该矩形的填充的前景色、背景色和渐变。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let fil = myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 50).Fill 　　　　fil.ForeColor.RGB = 128, 0, 0 　　　　fil.BackColor.RGB = 170, 170, 170 　　　　fil.TwoColorGradient(msoGradientHorizontal, 1) } ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将一条带图案的线条添加至 myDocument。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let lin = myDocument.Shapes.AddLine(10, 100, 250, 0).Line 　　　　lin.Weight = 6 　　　　lin.ForeColor.RGB = 0, 0, 255 　　　　lin.BackColor.RGB = 128, 0, 0 　　　　lin.Pattern = msoPatternDarkDownwardDiagonal }` |

#### **LineFormat.InsetPen**

该属性值为 **MsoTrue** 时，在指定形状的内部绘制线条。可读/写 **MsoTriState** 类型。

**语法**

**express.InsetPen**

*express*   一个代表 **LineFormat** 对象的变量。

**说明**

如果此属性试图在不支持插入笔绘图的任何 WPS Office自选图形中设置插入笔绘图，将发生错误。

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| msoCTrue 不适用于此属性。                     |
| msoFalse 默认值。不启用插入笔。               |
| msoTriStateMixed 不适用于此属性。             |
| msoTriStateToggle 不适用于此属性。            |
| msoTrue 启用插入笔。                          |

**示例**

下面的代码行在形状中启用插入笔。本示例假设当前演示文稿的第一张幻灯片中包含形状并且该形状支持插入笔绘图。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function DrawLinesInsideShape(){     Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).Line.InsetPen = msoTrue }` |

#### **LineFormat.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **LineFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test() {     let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes     let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame     addShp.TextRange.Text = "Test text"     addShp.Parent.Rotation = 45 } ` |

#### **LineFormat.Pattern**

设置或返回一个值，该值代表应用于指定线条的图案。**MsoPatternType** 类型，可读写。

**语法**

**express.Pattern**

*express*   一个代表 **LineFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 myDocument 中添加带有图案的线条。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let lin = myDocument.Shapes.AddLine(10, 100, 250, 0).Line 　　　　lin.Weight = 6 　　　　lin.ForeColor.RGB = 0, 0, 255 　　　　lin.BackColor.RGB = 128, 0, 0 　　　　lin.Pattern = msoPatternDarkDownwardDiagonal }` |

#### **LineFormat.Style**

返回或设置线条样式。可读/写 **MsoLineStyle** 类型。

**语法**

**express.Style**

*express*   一个代表 **LineFormat** 对象的变量。

**说明**

| MsoLineStyle 可以是下列 MsoLineStyle 常量之一。 |
| ----------------------------------------------- |
| msoLineSingle                                   |
| msoLineStyleMixed                               |
| msoLineThickBetweenThin                         |
| msoLineThickThin                                |
| msoLineThinThick                                |
| msoLineThinThin                                 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加粗的蓝色复合线。该复合线由一根粗线及其两侧的各一细线组成。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let lin = myDocument.Shapes.AddLine(10, 10, 250, 250).Line 　　　　lin.Style = msoLineThickBetweenThin 　　　　lin.Weight = 8 　　　　lin.ForeColor.RGB = 0, 0, 255 }` |

#### **LineFormat.Transparency**

该属性返回或设置指定填充、阴影或线条的透明度，可取 0.0（不透明）至 1.0（透明）之间的数值。**Single** 类型，可读/写。

**语法**

**express.Transparency**

*express*   一个代表 **LineFormat** 对象的变量。

**说明**

此属性的值仅影响均一颜色填充和线条的外观。对图案线条及渐变、图案、图片和纹理填充的外观无效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本例将 myDocument 中第三个形状的阴影设置为半透明的红色。如果形状原来无阴影，以下示例为它添加阴影。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let sha = myDocument.Shapes.Item(3).Shadow 　　　　sha.Visible = true 　　　　sha.ForeColor.RGB = 255, 0, 0 　　　　sha.Transparency = 0.5 }` |

#### **LineFormat.Visible**

返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。**MsoTriState** 类型。

**语法**

**express.Visible**

*express*   一个代表 **LineFormat** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| msoCTrue                                      |
| msoFalse                                      |
| msoTriStateMixed                              |
| msoTriStateToggle                             |
| msoTrue：指定对象或对象格式是可见的。         |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为当前演示文稿中第一张幻灯片的第三个形状设置阴影的水平偏移量和垂直偏移量，使其向右偏移 5 磅并向上偏移 3 磅。如果该形状原来没有阴影，本例将会为其添上阴影。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let sha = myDocument.Shapes.Item(3).Shadow 　　　　sha.Visible = msoTrue 　　　　sha.OffsetX = 5 　　　　sha.OffsetY = -3 }` |

#### **LineFormat.Weight**

以磅为单位返回或设置指定线条的粗细。可读/写 **Single** 类型。

**语法**

**express.Weight**

*express*   一个代表 **LineFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将一条两磅粗的绿色虚线添加到 myDocument。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let lin = myDocument.Shapes.AddLine(10, 10, 250, 250).Line 　　　　lin.DashStyle = msoLineDashDotDot 　　　　lin.ForeColor.RGB = 0, 255, 255 　　　　lin.Weight = 2 }` |

适用环境：web

适用平台：windows/linux