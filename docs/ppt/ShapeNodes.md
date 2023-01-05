**ShapeNodes**



指定的任意多边形中所有 **ShapeNode** 对象的集合。

**说明**

每个 **ShapeNode** 对象代表任意多边形中线段之间的一个节点或任意多边形曲线段的一个控制点。可以手动或使用 **BuildFreeform** 和 **ConvertToShape** 方法来创建任意多边形。

**方法**

|                                                              | 名称               | 说明                                                         |
| ------------------------------------------------------------ | ------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**         | 删除形状节点。                                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Insert**         | 将一个新线段插入到任意多边形的指定节点之后。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**           | 从指定集合中返回单个对象。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetEditingType** | 设置由 **Index** 指定的节点的编辑类型。如果节点是曲线段的控制点，则该方法将设置与之相邻的连接两条线段的节点的编辑类型。请注意，由于编辑类型的不同，该方法可能会影响相邻节点的位置。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetPosition**    | 设置由 **Index** 指定的节点的位置。请注意，由于节点的编辑类型的不同，该方法可能会影响相邻节点的位置。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetSegmentType** | 设置由 **Index** 指定的节点后面的线段的段类型。如果节点是曲线段的控制点，该方法将设置该曲线的段类型。请注意，该方法可能会因插入或删除相邻的节点而影响节点总数。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 **Application** 对象，该对象表示指定对象的创建者。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回指定集合中的对象数目。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回 **Long** 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。                                       |

**成员方法**

#### **ShapeNodes.Delete**

删除形状节点。

**语法**

**express.Delete(Index)**

*express*   一个代表 **ShapeNodes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Long**     | 指定要删除的节点。该节点后的段也将被删除。如果该节点是一条曲线的一个控制点，则该曲线及其所有节点都将被删除。 |

#### **ShapeNodes.Insert**

将一个新线段插入到任意多边形的指定节点之后。

**语法**

**express.Insert(Index, SegmentType, EditingType, X1, Y1, X2, Y2, X3, Y3)**

*express*   一个代表 **ShapeNodes** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型**       | **说明**                                                     |
| ------------- | ------------- | ------------------ | ------------------------------------------------------------ |
| *Index*       | 必选          | **Long**           | 要在其后插入新节点的节点。                                   |
| *SegmentType* | 必选          | **MsoSegmentType** | 要添加的线段的类型。                                         |
| *EditingType* | 必选          | **MsoEditingType** | 顶点的编辑属性。                                             |
| *X1*          | 必选          | **Single**         | 如果新线段的 EditingType 为 msoEditingAuto，则该参数指定从文档左上角到新线段终点的水平距离（以磅为单位）。如果新节点的 EditingType 为 msoEditingCorner，则该参数指定从文档左上角到新线段第一个控点的水平距离（以磅为单位）。 |
| *Y1*          | 必选          | **Single**         | 如果新线段的 EditingType 为 msoEditingAuto，则该参数指定从文档左上角到新线段终点的垂直距离（以磅为单位）。如果新节点的 EditingType 为 msoEditingCorner，则该参数指定从文档左上角到新线段第一个控点的垂直距离（以磅为单位）。 |
| *X2*          | 必选          | **Single**         | 如果新线段的 EditingType 为 msoEditingCorner，则该参数指定从文档左上角到新线段的第二个控制点的水平距离（以磅为单位）。如果新线段的 EditingType 为 msoEditingAuto，请不要指定该参数的值。 |
| *Y2*          | 必选          | **Single**         | 如果新线段的 EditingType 为 msoEditingCorner，则该参数指定从文档左上角到新线段的第二个控制点的垂直距离（以磅为单位）。如果新线段的 EditingType 为 msoEditingAuto，请不要指定该参数的值。 |
| *X3*          | 必选          | **Single**         | 如果新线段的 EditingType 为 msoEditingCorner，则该参数指定从文档左上角到新线段终点的水平距离（以磅为单位）。如果新线段的 EditingType 为 msoEditingAuto，请不要指定该参数的值。 |
| *Y3*          | 必选          | **Single**         | 如果新线段的 EditingType 为 msoEditingCorner，则该参数指定从文档左上角到新线段终点的垂直距离（以磅为单位）。如果新线段的 EditingType 为 msoEditingAuto，请不要指定该参数的值。 |

**说明**

| MsoSegmentType 可以是下列 MsoSegmentType 常量之一。 |
| --------------------------------------------------- |
| msoSegmentCurve                                     |
| msoSegmentLine                                      |

| MsoEditingType 可以是下列 MsoEditingType 常量之一（不能是 msoEditingSmooth 或 msoEditingSymmetric）。 |
| ------------------------------------------------------------ |
| msoEditingAuto                                               |
| msoEditingCorner                                             |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 myDocument 中第三个形状的第四个结点后添加一个带有一段曲线的平滑结点。第三个形状必须是至少有四个结点的任意多边形。*/ function test(){ 　　　　let myDocument = ActivePresentation.Slides.Item(1) 　　　　let mynodes = myDocument.Shapes.Item(3).Nodes 　　　　mynodes.Insert (4, msoSegmentCurve, msoEditingSmooth, 210, 100) }` |

#### **ShapeNodes.Item**

从指定集合中返回单个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **ShapeNodes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                               |
| -------- | ------------- | ------------ | -------------------------------------- |
| *Index*  | 必选          | **Variant**  | 集合中要返回的单个对象的名称或索引号。 |

**返回值**

ShapeNode

#### **ShapeNodes.SetEditingType**

设置由 **Index** 指定的节点的编辑类型。如果节点是曲线段的控制点，则该方法将设置与之相邻的连接两条线段的节点的编辑类型。请注意，由于编辑类型的不同，该方法可能会影响相邻节点的位置。

**语法**

**express.SetEditingType(Index, EditingType)**

*express*   一个代表 **ShapeNodes** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型**       | **说明**                 |
| ------------- | ------------- | ------------------ | ------------------------ |
| *Index*       | 必选          | **Long**           | 需要设置编辑类型的节点。 |
| *EditingType* | 必选          | **MsoEditingType** | 编辑类型。               |

**说明**

| MsoEditingType 可以是下列 MsoEditingType 常量之一。 |
| --------------------------------------------------- |
| msoEditingAuto                                      |
| msoEditingCorner                                    |
| msoEditingSmooth                                    |
| msoEditingSymmetric                                 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 myDocument 中第三个形状中的所有角结点更改为光滑结点。第三个形状必须是任意多边形。*/ function test(){ 　　　　let myDocument = ActivePresentation.Slides.Item(1) 　　　　let mynodes = myDocument.Shapes.Item(3).Nodes        for(let n = 1; n <= mynodes.Count; n++){  　　　　   if(mynodes.Item(n).EditingType == msoEditingCorner){      　　　　   mynodes.Item(n).SetEditingType (n, msoEditingSmooth)   　　　　  } 　　　　} }` |

#### **ShapeNodes.SetPosition**

设置由 **Index** 指定的节点的位置。请注意，由于节点的编辑类型的不同，该方法可能会影响相邻节点的位置。

**语法**

**express.SetPosition(Index, X1, Y1)**

*express*   一个代表 **ShapeNodes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                               |
| -------- | ------------- | ------------ | -------------------------------------- |
| *Index*  | 必选          | **Long**     | 要设置位置的顶点。                     |
| *X1, Y1* | 必选          | **Single**   | 新顶点相对于文档左上角的位置（点数）。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 myDocument 中第三个形状的第二个结点向右移 200 磅、向下移 300 磅。第三个形状必须是任意多边形。*/ function test(){ 　　　　let myDocument = ActivePresentation.Slides.Item(1) 　　　　let mynodes = myDocument.Shapes.Item(3).Nodes 　　　　let pointsArray = mynodes.Item(2).Points 　　　　let currXvalue = pointsArray(1, 1) 　　　　let currYvalue = pointsArray(1, 2) 　　　　mynodes.SetPosition (2, currXvalue + 200, currYvalue + 300) }` |

#### **ShapeNodes.SetSegmentType**

设置由 **Index** 指定的节点后面的线段的段类型。如果节点是曲线段的控制点，该方法将设置该曲线的段类型。请注意，该方法可能会因插入或删除相邻的节点而影响节点总数。

**语法**

**express.SetSegmentType(Index, SegmentType)**

*express*   一个代表 **ShapeNodes** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型**       | **说明**                 |
| ------------- | ------------- | ------------------ | ------------------------ |
| *Index*       | 必选          | **Long**           | 需要设置线段类型的顶点。 |
| *SegmentType* | 必选          | **MsoSegmentType** | 指明线段是直线还是曲线。 |

**说明**

| MsoSegmentType 可以是下列 MsoSegmentType 常量之一。 |
| --------------------------------------------------- |
| msoSegmentCurve                                     |
| msoSegmentLine                                      |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 myDocument 中第三个形状的所有直线段更改为曲线段。第三个形状必须是任意多边形。*/ function test(){ 　　　　let myDocument = ActivePresentation.Slides.Item(1) 　　　　let mynodes = myDocument.Shapes.Item(3).Nodes 　　　　let n = 1 　　　　while (n <= mynodes.Count){    　　　　 if(mynodes.Item(n).SegmentType == msoSegmentLine){  　　　　       mynodes.SetSegmentType (n, msoSegmentCurve) 　　　　    } 　　　　    n = n + 1 　　　　} }` |

**成员属性**

#### **ShapeNodes.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **ShapeNodes** 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function AddAndSave(pptPres){     pptPres.Slides.Add (1, 1)     pptPres.SaveAs (pptPres.Application.Path + "\\Added Slide") }` |

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ 　　　　let shpOle = ActivePresentation.Slides.Item(1).Shapes        for(let i = 1; i <= shpOle.Count; i++){    　　　　 if(shpOle.Item(i).Type == msoLinkedOLEObject){      　　　　   MsgBox (shpOle.Item(i).OLEFormat.Application.Name) 　　　　    } 　　　　} }` |

#### **ShapeNodes.Count**

返回指定集合中的对象数目。

**语法**

**express.Count**

*express*   一个代表 **ShapeNodes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭除当前窗口外的所有窗口。*/ function test(){ 　　　　let wins = Application.Windows     for(let i = 2; i <= wins.Count; i++){     　　　　wins.Item(2).Close() 　　　　} }` |

#### **ShapeNodes.Creator**

返回 **Long** 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。

**语法**

**express.Creator**

*express*   一个代表 **ShapeNodes** 对象的变量。

**说明**

**Creator** 属性用于 Macintosh 下的 WPS Office应用程序。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示关于 myObject 创建者的消息。*/ function test(){ 　　　　let myObject = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1) 　　　　if(myObject.Creator == "0x50575054"){ 　　　　    MsgBox ("This is a WPP object") 　　　　} 　　　　else{ 　　　　    MsgBox ("This is not a WPP object") 　　　　} }` |

#### **ShapeNodes.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **ShapeNodes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test(){ 　　　　let myShapes = ActivePresentation.Slides.Item(1).Shapes 　　　　let myshape = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame 　　　　myshape.TextRange.Text = "Test text" 　　　　myshape.Parent.Rotation = 45}` |

适用环境：web

适用平台：windows/linux