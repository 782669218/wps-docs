**ShapeNode**



代表用户定义的任意多边形中节点的几何图形以及几何图形编辑属性。

**说明**

节点包括任意多边形的线段之间的顶点以及曲线段的控制点。**ShapeNode** 对象是 **ShapeNodes** 集合的成员。**ShapeNodes** 集合包含任意多边形中的所有节点。

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 **Application** 对象，该对象表示指定对象的创建者。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回 **Long** 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EditingType** | 如果指定的节点是顶点，则此属性的返回值表示对节点所做的更改将如何影响与该节点相连的两条线段。如果该节点是曲线段的控制点，则此属性返回相邻顶点的编辑类型。只读**MsoEditingType** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Points**      | 返回一个 **Variant** 类型的值，该值以坐标对形式表示指定节点的位置。每个坐标以磅为单位表示。可以使用 **SetPosition** 方法设置该属性的值。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SegmentType** | 返回一个值，该值代表与指定节点相关联的线段是直线段还是曲线段。只读 **MsoSegmentType** 类型。 |

**成员属性**

#### **ShapeNode.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **ShapeNode** 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function AddAndSave(pptPres){     pptPres.Slides.Add (1, 1)     pptPres.SaveAs (pptPres.Application.Path + "\\Added Slide") }` |

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ 　　　　let shpOle = ActivePresentation.Slides.Item(1).Shapes        for(let i = 1; i <= shpOle.Count; i++){     　　　　if(shpOle.Item(i).Type == msoLinkedOLEObject){        　　　　 MsgBox (shpOle.Item(i).OLEFormat.Application.Name) 　　　　    } 　　　　} }` |

#### **ShapeNode.Creator**

返回 **Long** 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。

**语法**

**express.Creator**

*express*   一个代表 **ShapeNode** 对象的变量。

**说明**

**Creator** 属性用于 Macintosh 下的 WPS Office应用程序。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示关于 myObject 创建者的消息。*/ function test(){ 　　　　let myObject = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1) 　　　　if(myObject.Creator == "&h50575054"){ 　　　　    MsgBox ("This is aWPP object") 　　　　} 　　　　else{ 　　　　    MsgBox ("This is not aWPP object") 　　　　} }` |

#### **ShapeNode.EditingType**

如果指定的节点是顶点，则此属性的返回值表示对节点所做的更改将如何影响与该节点相连的两条线段。如果该节点是曲线段的控制点，则此属性返回相邻顶点的编辑类型。只读**MsoEditingType** 类型。

**语法**

**express.EditingType**

*express*   一个代表 **ShapeNode** 对象的变量。

**说明**

此属性为只读。可用 **SetEditingType** 方法设置此属性的值。

| MsoEditingType 可以是下列 MsoEditingType 常量之一。 |
| --------------------------------------------------- |
| msoEditingAuto                                      |
| msoEditingCorner                                    |
| msoEditingSmooth                                    |
| msoEditingSymmetric                                 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 myDocument 中第三个形状中的所有角结点更改为光滑结点。第三个形状必须是任意多边形。*/ function test(){ 　　　　let myDocument = ActivePresentation.Slides.Item(1) 　　　　let mynodes = myDocument.Shapes.Item(3).Nodes     for(let n = 1; n <= mynodes.Count; n++){    　　　　 if(mynodes.Item(n).EditingType == msoEditingCorner){   　　　　      mynodes.Item(n).SetEditingType (n, msoEditingSmooth) 　　　　    } 　　　　} }` |

#### **ShapeNode.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **ShapeNode** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test(){ 　　　　let myShapes = ActivePresentation.Slides.Item(1).Shapes 　　　　let myshape = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame 　　　　myshape.TextRange.Text = "Test text" 　　　　myshape.Parent.Rotation = 45 }` |

#### **ShapeNode.Points**

返回一个 **Variant** 类型的值，该值以坐标对形式表示指定节点的位置。每个坐标以磅为单位表示。可以使用 **SetPosition** 方法设置该属性的值。只读。

**语法**

**express.Points**

*express*   一个代表 **ShapeNode** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将当前演示文稿中第三个形状的第二个结点向右移 200 磅，并向下移 300 磅。第三个形状必须是任意多边形。*/ function test(){ 　　　　let mynodes = ActivePresentation.Slides.Item(1).Shapes.Item(3).Nodes 　　　　let pointsArray = mynodes.Item(2).Points 　　　　let currXvalue = pointsArray(1, 1) 　　　　let currYvalue = pointsArray(1, 2) 　　　　mynodes.SetPosition (2, currXvalue + 200, currYvalue + 300) }` |

#### **ShapeNode.SegmentType**

返回一个值，该值代表与指定节点相关联的线段是直线段还是曲线段。只读 **MsoSegmentType** 类型。

**语法**

**express.SegmentType**

*express*   一个代表 **ShapeNode** 对象的变量。

**说明**

此属性为只读。可使用 **SetSegmentType** 方法设置此属性的值。

| MsoSegmentType 可以是下列 MsoSegmentType 常量之一。          |
| ------------------------------------------------------------ |
| msoSegmentCurve 如果指定节点是曲线段的控制点，则 SegmentType 属性返回该值。 |
| msoSegmentLine                                               |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 myDocument 中第三个形状的所有直线段更改为曲线段。第三个形状必须是任意多边形。*/ function test(){ 　　　　let myDocument = ActivePresentation.Slides.Item(1) 　　　　let mynodes = myDocument.Shapes.Item(3).Nodes 　　　　let n = 1 　　　　while (n <= mynodes.Count){    　　　　 if(mynodes.Item(n).SegmentType == msoSegmentLine){ 　　　　        mynodes.SetSegmentType (n, msoSegmentCurve) 　　　　    } 　　　　    n = n + 1 　　　　} }` |

适用环境：web

适用平台：windows/linux