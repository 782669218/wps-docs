**Adjustments**



包含指定自选图形、艺术字对象或连接符的调整值的集合。

**说明**

每个调整值代表了调整控点可以调整的方向。因为某些调整控点有两个调整方向（例如，某些调整控点可以在水平和垂直两个方向调整），所以形状的调整值数量可以大于调整控点的数量。一个形状最多可有八个调整值。

使用 **Adjustments** 属性返回 **Adjustments** 对象。使用 **Adjustments**(*index*) 返回单个调整值，其中 *index* 是该调整值的索引号。

不同的形状有不同数量的调整值，不同的调整值以不同的方式改变形状的几何外形，且不同的调整值有不同的有效值范围。例如，以下示例显示右箭头标注的四个调整值分别如何影响该标注几何外形的定义。

| ![img]() 注释                                                |
| ------------------------------------------------------------ |
| 因为每个可调整的形状都有一组不同的调整值，所以检验某一特定形状的调整行为的建议方法是：先手动创建该形状的一个实例，再打开宏记录器并调整该形状，然后查看所记录的代码。 |

下表概述了不同类型的调整所具有的有效的调整值范围。多数情况下，如果指定的值超过了有效值范围，将给调整分配最接近该值的有效值。

| 调整类型           | 有效值                                                       |
| ------------------ | ------------------------------------------------------------ |
| 线性（水平或垂直） | 通常 0.0 值代表形状的左边界或上边界，而 1.0 值代表形状的右边界或下边界。有效值对应于有效的手动调整。例如，如果只能将调整控点手动拖动形状的一半宽度，则相应的调整值最大为 0.5。对于象连接符和标注这样的形状，0.0 和 1.0 值代表由它们的起始和终止点定义的矩形界限，此时负值和大于 1.0 的值是有效的。 |
| 射线               | 调整值 1.0 对应于形状宽度。最大值为 0.5，或形状宽度的一半。  |
| 角                 | 值以度表示。如果指定的值超过了 -180 到 180 的范围，则将其折算为该范围内的值。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 Application 对象，该对象表示指定对象的创建者。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回指定集合中的对象数目。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回 Long 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Item**        | 返回或设置由 Index 参数指定的调整值。对于线性调整，0.0 调整值通常对应于形状的左边缘或上边缘，1.0 调整值通常对应于形状的右边缘或下边缘。但对某些形状，调整值可以超过形状边界。对于径向调整，1.0 调整值对应于形状宽度。对于角度调整，调整值用度数指定。Item 属性仅应用于含有调整的形状。可读/写。Single 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。                                       |

**成员属性**

#### **Adjustments.Application**

返回一个 Application 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **Adjustments** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/ function AddAndSave(pptPres) {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/ function test(){     for(let i=1;i <= Application.ActivePresentation.Slides.Item(1).Shapes.Count;i++) {         if(Application.ActivePresentation.Slides.Item(1).Shapes.Item(i).Type == msoLinkedOLEObject) {          　　　alert(Application.ActivePresentation.Slides.Item(1).Shapes.Item(i).OLEFormat.Application.Name)       　}     } }` |

#### **Adjustments.Count**

返回指定集合中的对象数目。

**语法**

**express.Count**

*express*   一个代表 **Adjustments** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭第2个窗口。*/ function test(){ 　　let Win = Application.Windows     for(let i=2;i <= Win.Count;i++) { 　　　　 Win.Item(2).Close() 　　} }` |

#### **Adjustments.Creator**

返回 Long 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。

**语法**

**express.Creator**

*express*   一个代表 **Adjustments** 对象的变量。

**说明**

**Creator** 属性用于 Macintosh 下的 WPS Office应用程序。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示关于 myObject 创建者的消息。*/ function test(){ 　　　　let myObject = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1) 　　　　if(myObject.Creator == parseInt(50575054,16)) { 　　　　    alert("This is a WPP object") 　　　　} 　　　　else { 　　　　    alert("This is not a WPP object") 　　　　}}` |

#### **Adjustments.Item**

返回或设置由 Index 参数指定的调整值。对于线性调整，0.0 调整值通常对应于形状的左边缘或上边缘，1.0 调整值通常对应于形状的右边缘或下边缘。但对某些形状，调整值可以超过形状边界。对于径向调整，1.0 调整值对应于形状宽度。对于角度调整，调整值用度数指定。Item 属性仅应用于含有调整的形状。可读/写。Single 类型。

**语法**

**express.Item**

*express*   一个代表 **Adjustments** 对象的变量。

**说明**

自选图形、连接符和艺术字对象最多可有八个调整。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将两个十字形添加到 myDocument。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let myShape = myDocument.Shapes 　　　　myShape.AddShape(msoShapeCross, 10, 10, 100, 100).Adjustments.Item(1)  　　　　myShape.AddShape(msoShapeCross, 150, 10, 100, 100).Adjustments.Item(1) }` |

#### **Adjustments.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Adjustments** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test(){     let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes     let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame     addShp.TextRange.Text = "Test text"     addShp.Parent.Rotation = 45 }` |

适用环境：web

适用平台：windows/linux