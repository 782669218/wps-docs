#### **Shapes**



**Shape** 对象的集合，这些对象代表某个文档中的所有形状，或某个文档的所有页眉和页脚中的所有形状。每个 **Shape** 对象代表绘制层的一个对象，如自选图形、任意多边形、OLE 对象或图片。

**方法**

|                                                              | 名称              | 说明                                                         |
| ------------------------------------------------------------ | ----------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddCanvas**     | 在文档中添加绘图画布。返回代表该画布的 **Shape** 对象并将其添加到 **Shapes** 集合。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddChart**      | 将指定类型的图表作为形状插入到活动文档中，并且在ET 中打开包含 WPS 用于创建该图表的默认数据的工作表。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddCurve**      | 返回一个 **Shape** 对象，该对象代表绘图画布上的贝赛尔曲线。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddLabel**      | 在绘图画布上添加一个文本标签。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddLine**       | 在绘图画布上添加一条直线。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddOLEControl** | 创建一个 ActiveX 控件（以前称为 OLE 控件）。返回代表新 ActiveX 控件的 **InlineShape** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddOLEObject**  | 创建一个 OLE 对象。返回代表新 OLE 对象的 **InlineShape** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddPicture**    | 在绘图画布上添加一幅图片。返回一个 **Shape** 对象，该对象代表图片并将其添加至 **CanvasShapes** 集合。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddPolyline**   | 在绘图画布上添加一个开放或封闭的多边形。返回一个代表该多边形的 **Shape** 对象，并将其添加到 **CanvasShapes** 集合。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddShape**      | 向某文档中添加一个自选图形。返回代表该自选图形的 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象并将其添加到 **Shapes** 集合。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddSmartArt**   | 将指定 SmartArt 图形插入活动文档。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddTextbox**    | 在绘图画布上添加一个文本框。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddTextEffect** | 在绘图画布上添加一个“艺术字”图形。返回一个 **Shape** 对象，该对象代表“艺术字”，并将其添加至 **CanvasShapes** 集合。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **BuildFreeform** | 建立任意多边形对象。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**          | 返回集合中的单个 **Shape** 对象。                            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Range**         | 返回代表范围内的形状的 **ShapeRange** 对象。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectAll**     | 选择形状集合中的所有形状。                                   |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AddCallout**  | 在绘图画布上添加无框线的标注。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 类型的值，该值代表集合中图形的数量。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **Shapes** 对象的父对象。 |

**成员方法**

#### **Shapes.AddCanvas**

在文档中添加绘图画布。返回代表该画布的 **Shape** 对象并将其添加到 **Shapes** 集合。

**语法**

**express.AddCanvas(Left, Top, Width, Height, Anchor)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Left*   | 必选          | **Single**   | 画布左侧边缘相对于锁定标记的位置，以磅为单位。               |
| *Top*    | 必选          | **Single**   | 画布上部边缘相对于锁定标记的位置，以磅为单位。               |
| *Width*  | 必选          | **Single**   | 画布的宽度，以磅为单位。                                     |
| *Height* | 必选          | **Single**   | 画布的高度，以磅为单位。                                     |
| *Anchor* | 可选          | **Variant**  | 代表画布所绑定的文本的 Range 对象。如果指定 Anchor，则锁定标记将出现在锁定区域第一段的开头。如果省略该参数，将自动选定锁定区域，而画布将相对于页面的上边缘和左边缘进行定位。 |

**返回值**

Shape

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在新文档中添加画布，然后在画布上添加两个图形，并设置填充和线条属性。*/ function AddInlineCanvas(){     //Add a drawing canvas to the new document     let shpCanvas = Documents.Add().Shapes.AddCanvas(150, 150, 70, 70)     shpCanvas.WrapFormat.Type = wdWrapInline     //Add shapes to drawing canvas     let canvasitems = shpCanvas.CanvasItems     canvasitems.AddShape(msoShapeHeart, 10, 10, 50, 60)     canvasitems.AddLine(0, 0, 70, 70)     let shpcanvas = shpCanvas     shpcanvas.CanvasItems.Item(1).Fill.ForeColor.RGB = 255, 0, 0     shpcanvas.CanvasItems.Item(2).Line.EndArrowheadStyle = msoArrowheadTriangle }` |

#### **Shapes.AddChart**

将指定类型的图表作为形状插入到活动文档中，并且在ET 中打开包含 WPS 用于创建该图表的默认数据的工作表。

**语法**

**express.AddChart(Type, Left, Top, Width, Height, Anchor)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**    | **说明**                                                     |
| -------- | ------------- | --------------- | ------------------------------------------------------------ |
| *Type*   | 可选          | **XlChartType** | 指定要创建的图表的类型。                                     |
| *Left*   | 可选          | **Variant**     | 图表左边缘相对于定位标记的位置（以磅为单位）。               |
| *Top*    | 可选          | **Variant**     | 图表上边缘相对于定位标记的位置（以磅为单位）。               |
| *Width*  | 可选          | **Variant**     | 图表的宽度（以磅为单位）。                                   |
| *Height* | 可选          | **Variant**     | 图表的高度（以磅为单位）。                                   |
| *Anchor* | 可选          | **Variant**     | 代表图表所绑定文本的 Range 对象。如果指定 Anchor，则定位标记将出现在定位区域中第一段的开头。如果省略该参数，则自动选择定位区域，并且相对于页面的上边缘和左边缘放置图表。 |

**返回值**

Shape

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*创建一个新的三维柱形图，然后将其添加到活动文档中。*/ ActiveDocument.Shapes.AddChart(xl3DColumn)` |

#### **Shapes.AddCurve**

返回一个 **Shape** 对象，该对象代表绘图画布上的贝赛尔曲线。

**语法**

**express.AddCurve(SafeArrayOfPoints)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称**            | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *SafeArrayOfPoints* | 必选          | **Variant**  | 指定该曲线的顶点和控制点的坐标对?（坐标对：一对值，表示两维数组中存储的点的 x 和 y 坐标，该数组中包含许多点的坐标。）数组。指定的第一个点是起始顶点，随后指定的两个点是第一段贝塞尔曲线的控制点。如果该曲线还有其他段，则每段都应该指定一个顶点和两个控制点。指定的最后一个点是该曲线的终止顶点。请注意，须始终指定 3n + 1 个点，此处 n 为曲线的段数。 |

**返回值**

Shape

**说明**

------

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在新绘图画布上添加一条贝赛尔曲线。*/ function CanvasBezier(){     let sngArray = [[0,0],[0,0],[0,0],[0,0],[0,0],[0,0],[0,0]]     let docNew = Documents.Add()     //Create a new drawing canvas     let shpCanvas = docNew.Shapes.AddCanvas(100, 100, 300, 50)     sngArray[0][0] = 0     sngArray[0][1] = 0     sngArray[1][0] = 50     sngArray[1][1] = 50     sngArray[2][0] = 100     sngArray[2][1] = 0     sngArray[3][0] = 150     sngArray[3][1] = 50     sngArray[4][0] = 200     sngArray[4][1] = 0     sngArray[5][0] = 250     sngArray[5][1] = 50     sngArray[6][0] = 300     sngArray[6][1] = 0     //Add Bezier curve to drawing canvas     shpCanvas.CanvasItems.AddCurve(sngArray) }` |

#### **Shapes.AddLabel**

在绘图画布上添加一个文本标签。

**语法**

**express.AddLabel(Orientation, Left, Top, Width, Height)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型**           | **说明**                                             |
| ------------- | ------------- | ---------------------- | ---------------------------------------------------- |
| *Orientation* | 必选          | **MsoTextOrientation** | 文本的方向。                                         |
| *Left*        | 必选          | **Single**             | 标签左边缘相对于绘图画布左边缘的位置（以磅为单位）。 |
| *Top*         | 必选          | **Single**             | 标签上边缘相对于绘图画布上边缘的位置（以磅为单位）。 |
| *Width*       | 必选          | **Single**             | 标签的宽度（以磅为单位）。                           |
| *Height*      | 必选          | **Single**             | 标签的高度（以磅为单位）。                           |

**返回值**

Shapes

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档的新绘图画布上添加一个蓝色的文本标签，该标签带有文字“Hello World”。*/ function NewCanvasTextLabel(){     //Add a drawing canvas to the active document     let shpCanvas = ActiveDocument.Shapes.AddCanvas(100, 75, 150, 200)          //Add a label to the drawing canvas     let shpLabel = shpCanvas.CanvasItems.AddLabel(msoTextOrientationHorizontal, 15, 15, 100, 100)          //Fill the label textbox with a color,     //add text to the label and format it     let fill = shpLabel.Fill         fill.BackColor.RGB = 0, 0, 192         //Make the fill visible         fill.Visible = msoTrue     let newtextrange = shpLabel.TextFrame.TextRange         newtexttange.Text = "Hello World."         newtexttange.Bold = true         newtexttange.Font.Name = "Tahoma" }` |

#### **Shapes.AddLine**

在绘图画布上添加一条直线。

**语法**

**express.AddLine(BeginX, BeginY, EndX, EndY)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                         |
| -------- | ------------- | ------------ | ------------------------------------------------ |
| *BeginX* | 必选          | **Single**   | 直线起点相对于绘图画布的水平位置（以磅为单位）。 |
| *BeginY* | 必选          | **Single**   | 直线起点相对于绘图画布的垂直位置（以磅为单位）。 |
| *EndX*   | 必选          | **Single**   | 直线终点相对于绘图画布的水平位置（以磅为单位）。 |
| *EndY*   | 必选          | **Single**   | 直线终点相对于绘图画布的垂直位置（以磅为单位）。 |

**返回值**

Shape

**说明**

若要创建带箭头的直线，请使用 **Line** 属性来设置线条格式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在新绘图画布上添加一条紫色带箭头的直线。*/ function NewCanvasLine(){     //Add new drawing canvas to the active document     let shpCanvas = ActiveDocument.Shapes.AddCanvas(100, 75, 150, 200)      //Add a line to the drawing canvas     let shpLine = shpCanvas.CanvasItems.AddLine(25, 25, 150, 150)      //Add an arrow to the line and sets the color to purple     let line = shpLine.Line         line.BeginArrowheadStyle = msoArrowheadDiamond         line.BeginArrowheadWidth = msoArrowheadWide         line.ForeColor.RGB = 150, 0, 255 }` |

#### **Shapes.AddOLEControl**

创建一个 ActiveX 控件（以前称为 OLE 控件）。返回代表新 ActiveX 控件的 **InlineShape** 对象。

**语法**

**express.AddOLEControl(ClassType, Range)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *ClassType* | 可选          | **Variant**  | 用于要创建的 ActiveX 控件的编程标识符。                      |
| *Range*     | 可选          | **Variant**  | 一个范围，ActiveX 控件将放置在该范围的文本中。如果该范围未折叠，则 ActiveX 控件将替换该范围。如果省略该参数，则自动放置 ActiveX 控件。 |

**返回值**

InlineShape

**说明**

在 WPS 中，将 ActiveX 控件表示为 **Shape** 对象或 **InlineShape** 对象。要修改 ActiveX 控件的属性，可以对指定的形状或内嵌形状使用 **OLEFormat** 对象的 **Object** 属性。

有关可用的 ActiveX 控件类类型的信息，请参阅 OLE 编程标识符。

#### **Shapes.AddOLEObject**

创建一个 OLE 对象。返回代表新 OLE 对象的 **InlineShape** 对象。

**语法**

**express.AddOLEObject(ClassType, FileName, LinkToFile, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Range)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *ClassType*     | 可选          | **Variant**  | 用于激活指定 OLE 对象的应用程序的名称。                      |
| *FileName*      | 可选          | **Variant**  | 创建对象的文件。如果省略该参数，则使用当前文件夹。必须指定对象的 ClassType 或 FileName 参数，但不能同时指定两者。 |
| *LinkToFile*    | 可选          | **Variant**  | 如果该参数值为 True，则将 OLE 对象链接到创建该对象的文件。如果该参数值为 False，则 OLE 对象成为该文件的独立副本。如果该参数值为 ClassType 指定了值，则 LinkToFile 参数必须为 False。默认值为 False。 |
| *DisplayAsIcon* | 可选          | **Variant**  | 如果该参数值为 True，则将 OLE 对象显示为图标。默认值为 False。 |
| *IconFileName*  | 可选          | **Variant**  | 包含将要显示的图标的文件。                                   |
| *IconIndex*     | 可选          | **Variant**  | IconFileName 内的图标的索引号。当选中“显示为图标”复选框时，指定文件中图标的顺序对应于“更改图标”对话框中图标显示的顺序。文件中的第一个图标的索引号为 0（零）。如果 IconFileName 中不存在指定索引号的图标，则使用索引号为 1 的图标（文件中的第二个图标）。默认值为 0（零）。 |
| *IconLabel*     | 可选          | **Variant**  | 显示在图标下面的标签（题注）。                               |
| *Range*         | 可选          | **Variant**  | 一个范围，OLE 对象将放置在该范围的文本中。除非范围是折叠的，否则 OLE 对象将替换该范围。如果省略此参数，则自动放置对象。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动文档中添加一个新的浮动位图映射。该位图链接到另一个文件。*/ ActiveDocument.Shapes.AddOLEObject(null,"c:\\my documents\\MyDrawing.bmp", true)` |

#### **Shapes.AddPicture**

在绘图画布上添加一幅图片。返回一个 **Shape** 对象，该对象代表图片并将其添加至 **CanvasShapes** 集合。

**语法**

**express.AddPicture(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称**           | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *FileName*         | 必选          | **String**   | 图片的路径和文件名。                                         |
| *LinkToFile*       | 可选          | **Variant**  | 如果该参数值为 **True**，则将图片链接到创建它的文件。如果该参数值为 **False**，则将图片作为该文件的独立副本。默认值为 False。 |
| *SaveWithDocument* | 可选          | **Variant**  | 如果该参数值为 **True**，则将链接的图片与文档一起保存。默认值为 **False**。 |
| *Left*             | 可选          | **Variant**  | 新图片的左边缘相对于绘图画布的位置，以磅为单位。             |
| *Top*              | 可选          | **Variant**  | 新图片的上边缘相对于绘图画布的位置，以磅为单位。             |
| *Width*            | 可选          | **Variant**  | 图片的宽度，以磅为单位。                                     |
| *Height*           | 可选          | **Variant**  | 图片的高度，以磅为单位。                                     |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动文档中新创建的绘图画布上添加一幅图片*/ function NewCanvasPicture(){     //Add a drawing canvas to the active document     let shpCanvas = Application.ActiveDocument.Shapes.AddCanvas(100, 75, 200, 300)      //Add a graphic to the drawing canvas     shpCanvas.CanvasItems.AddPicture("C:\\Program Files\\Microsoft Office\\" + "Office\\Bitmaps\\Styles\\stone.bmp", false, true) }` |

#### **Shapes.AddPolyline**

在绘图画布上添加一个开放或封闭的多边形。返回一个代表该多边形的 **Shape** 对象，并将其添加到 **CanvasShapes** 集合。

**语法**

**express.AddPolyline(SafeArrayOfPoints)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称**            | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *SafeArrayOfPoints* | 必选          | **Variant**  | 指定连续线段图形的顶点的坐标对?（坐标对：一对值，表示两维数组中存储的点的 x 和 y 坐标，该数组中包含许多点的坐标。）数组。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在新绘图画布上创建一个 V 形开放多边形。*/ function NewCanvasPolyline(){     let sngArray = [[0,0],[0,0],[0,0]]     //Creates a new document and adds a drawing canvas     let docNew = Documents.Add()     let shpCanvas = docNew.Shapes.AddCanvas(100, 75, 200, 300)      //Sets the coordinates of the array     sngArray[0][0] = 100     sngArray[0][1] = 75     sngArray[1][0] = 150     sngArray[1][1] = 100     sngArray[2][0] = 100     sngArray[2][1] = 125 	     //Adds a V-shaped open polyline to the drawing canvas     shpCanvas.CanvasItems.AddPolyline(sngArray) }` |

#### **Shapes.AddShape**

向某文档中添加一个自选图形。返回代表该自选图形的 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象并将其添加到 **Shapes** 集合。

**语法**

**express.AddShape(Type, Left, Top, Width, Height)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                 |
| -------- | ------------- | ------------ | -------------------------------------------------------- |
| *Type*   | 必选          | **Long**     | 要返回的图形类型。可以是任何 **MsoAutoShapeType** 常量。 |
| *Left*   | 必选          | **Single**   | 自选图形左边缘的位置（以磅为单位）。                     |
| *Top*    | 必选          | **Single**   | 自选图形上边缘的位置（以磅为单位）。                     |
| *Width*  | 必选          | **Single**   | 自选图形的宽度（以磅为单位）。                           |
| *Height* | 必选          | **Single**   | 自选图形的高度（以磅为单位）。                           |

**说明**

要改变所添加的自选图形类型，可设置其 **AutoShapeType** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*向活动文档中添加一个矩形*/ Application.ActiveDocument.Shapes.AddShape(1,100, 100, 100, 100)` |

#### **Shapes.AddSmartArt**

将指定 SmartArt 图形插入活动文档。

**语法**

**express.AddSmartArt(Layout, Left, Top, Width, Height, Anchor)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**         | **说明**                                                     |
| -------- | ------------- | -------------------- | ------------------------------------------------------------ |
| *Layout* | 必选          | **[SMARTARTLAYOUT]** | 一个指定 SmartArt 图形的布局的 SmartArtLayout 对象。         |
| *Left*   | 可选          | **Variant**          | 从幻灯片左边缘到 SmartArt 图形左边缘之间的距离（以磅为单位）。 |
| *Top*    | 可选          | **Variant**          | 从幻灯片上边缘到 SmartArt 图形上边缘之间的距离（以磅为单位）。 |
| *Width*  | 可选          | **Variant**          | SmartArt 图形的宽度。                                        |
| *Height* | 可选          | **Variant**          | SmartArt 图形的高度。                                        |
| *Anchor* | 可选          | **Variant**          | 一个代表 SmartArt 所绑定的文本的 Range 对象。如果指定 Anchor，则定位标记将出现在定位区域第一段的开头。如果省略该参数，将自动选定定位区域，并且相对于页面的上边缘和左边缘放置 SmartArt 图形。 |

**返回值**

Shape

#### **Shapes.AddTextbox**

在绘图画布上添加一个文本框。

**语法**

**express.AddTextbox(Orientation, Left, Top, Width, Height)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型**           | **说明**                                                     |
| ------------- | ------------- | ---------------------- | ------------------------------------------------------------ |
| *Orientation* | 必选          | **MsoTextOrientation** | 文本的方向。由于选择或安装的语言支持（例如美国英语）不同，有些常量可能无法使用。 |
| *Left*        | 必选          | **Single**             | 文本框左边缘的位置（以磅为单位）。                           |
| *Top*         | 必选          | **Single**             | 文本框上边缘的位置（以磅为单位）。                           |
| *Width*       | 必选          | **Single**             | 文本框的宽度（以磅为单位）。                                 |
| *Height*      | 必选          | **Single**             | 文本框的高度（以磅为单位）。                                 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在新文档中的绘图画布上添加一个文本框。*/ function NewCanvasTextbox(){     //Create a new document and add a drawing canvas     let docNew = Application.Documents.Add()     let shpCanvas = docNew.Shapes.AddCanvas(100, 75, 150, 200)      //Add a text box to the drawing canvas     shpCanvas.CanvasItems.AddTextbox(msoTextOrientationHorizontal, 1, 1, 100, 100) }` |

#### **Shapes.AddTextEffect**

在绘图画布上添加一个“艺术字”图形。返回一个 **Shape** 对象，该对象代表“艺术字”，并将其添加至 **CanvasShapes** 集合。

**语法**

**express.AddTextEffect(PresetTextEffect, Text, FontName, FontSize, FontBold, FontItalic, Left, Top)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称**           | **必选/可选** | **数据类型**            | **说明**                                                     |
| ------------------ | ------------- | ----------------------- | ------------------------------------------------------------ |
| *PresetTextEffect* | 必选          | **MsoPresetTextEffect** | 预设的文本效果。MsoPresetTextEffect 常量的值对应于“‘艺术字’库”对话框中所列的格式（按从上到下、从左到右的顺序编号）。 |
| *Text*             | 必选          | **String**              | 艺术字中的文字。                                             |
| *FontName*         | 必选          | **String**              | 艺术字中所用字体的名称。                                     |
| *FontSize*         | 必选          | **Single**              | 艺术字中所用字体的大小（以磅为单位）。                       |
| *FontBold*         | 必选          | **MsoTriState**         | 如果为 MsoTrue，则对艺术字字体应用加粗格式。                 |
| *FontItalic*       | 必选          | **MsoTriState**         | 如果为 MsoTrue，则对艺术字字体应用倾斜格式。                 |
| *Left*             | 必选          | **Single**              | 艺术字形状左边缘相对于绘图画布左边缘的位置（以磅为单位）。   |
| *Top*              | 必选          | **Single**              | 艺术字形状上边缘相对于绘图画布上边缘的位置（以磅为单位）。   |

**说明**

向文档添加艺术字对象时，该对象的高度和宽度将自动根据所指定的文字的大小和数量来设置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在新文档中添加一个绘图画布，并在绘图画布中插入包含文字“Hello, World”的“艺术字”图形。*/ function test(){     //Create a new document and add a drawing canvas     let docNew = Documents.Add()     let shpCanvas = docNew.Shapes.AddCanvas(100, 100, 150, 50)      //Add WordArt shape to the drawing canvas     shpCanvas.CanvasItems.AddTextEffect(msoTextEffect20, "Hello, World", "Tahoma", 15, msoTrue, msoFalse, 120, 120) }` |

#### **Shapes.BuildFreeform**

建立任意多边形对象。

**语法**

**express.BuildFreeform(EditingType, X1, Y1)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型**       | **说明**                                                   |
| ------------- | ------------- | ------------------ | ---------------------------------------------------------- |
| *EditingType* | 必选          | **MsoEditingType** | 第一个顶点的编辑属性。                                     |
| *X1*          | 必选          | **Single**         | 任意多边形第一个顶点相对于文档左边缘的位置（以磅为单位）。 |
| *Y1*          | 必选          | **Single**         | 任意多边形第一个顶点相对于文档上边缘的位置（以磅为单位）。 |

**返回值**

FreeformBuilder

**说明**

使用 **AddNodes** 方法将段添加到任意多边形。向该多边形添加至少一个段后，可使用 **ConvertToShape** 方法将 **FreeformBuilder** 对象转换为具有在 **FreeformBuilder** 对象中定义的几何形状描述的 **Shape** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将一个具有五个顶点的任意多边形添加到活动文档中。*/ function test() { let docActive = ActiveDocument.Shapes.BuildFreeform(msoEditingCorner, 360, 200)     docActive.AddNodes(msoSegmentCurve, msoEditingCorner, 380, 230, 400, 250, 450, 300)     docActive.AddNodes(msoSegmentCurve, msoEditingAuto, 480, 200)     docActive.AddNodes(msoSegmentLine, msoEditingAuto, 480, 400)     docActive.AddNodes(msoSegmentLine, msoEditingAuto, 360, 200)     docActive.ConvertToShape() }` |

#### **Shapes.Item**

返回集合中的单个 **Shape** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Variant**  | 要返回的单个对象。可以是代表序号位置的 **Long** 类型值，或代表单个对象名称的 **String 类型值。** |

#### **Shapes.Range**

返回代表范围内的形状的 **ShapeRange** 对象。

**语法**

**express.Range(Index)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Variant**  | 指定要在指定区域内包含哪些形状。可以是指定 Shapes 集合内某个形状的索引号的整数，也可以是指定形状名称的字符串，或者是包含整数或字符串的数组。 |

**返回值**

ShapeRange

**说明**

**Shape** 对象总是与定位的区域位于同一页上。

| 注释                                                         |
| ------------------------------------------------------------ |
| 对 **Shape** 对象进行的大部分操作也适用于只包含一个形状的 **ShapeRange** 对象。有些操作如果用于包含多个形状的 **ShapeRange** 对象，则会产生错误。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一个形状的填充前景色设置为紫色。*/ function ShRange(){     let range = ActiveDocument.Shapes.Range(1).Fill         range.ForeColor.RGB = 255, 0, 255         range.Visible = msoTrue }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为活动文档中变量形状应用阴影。*/ function ShpRange2(strShpName){     ActiveDocument.Shapes.Range(strShpName).Shadow.Type = msoShadow6 }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*要调用前一个子程序，请在标准代码模块中输入以下代码。*/ function CallShpRange2(){     let shpArrow = ActiveDocument.Shapes.AddShape(msoShapeLeftArrow, 200, 400, 50, 75)     shpArrow.Name = "myShape"     let strName = shpArrow.Name     ShpRange2(strName) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例选择活动文档的第一和第三个形状。*/ function SelectShapeRange(){     ActiveDocument.Shapes.Range(Array(1, 3)).Select() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例选择并删除活动文档的第一个形状中的形状。本示例假定第一个形状是画布形状。*/ function CanvasShapeRange(){     let rngCanvasShapes = ActiveDocument.Shapes.Item(1).CanvasItems.Range(1)         rngCanvasShapes.Select()         rngCanvasShapes.Delete() }` |

#### **Shapes.SelectAll**

选择形状集合中的所有形状。

**语法**

**express.SelectAll()**

*express*   一个代表 **Shapes** 对象的变量。

**说明**

此方法不能选择 **InlineShape** 对象。不能使用此方法选择多个画布。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例选择活动文档中的所有形状。*/ function SelectAllShapes(){     ActiveDocument.Shapes.SelectAll() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例选择活动文档的页眉和页脚中的所有形状，并给每一个形状添加红色阴影。*/ function SelectAllHeaderShapes(){     let selectall = ActiveDocument.ActiveWindow         selectall.View.Type = wdPrintView         selectall.ActivePane.View.SeekView = wdSeekCurrentPageHeader     ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterPrimary).Shapes.SelectAll()     let shadow = Selection.ShapeRange.Shadow         shadow.Type = msoShadow10         shadow.ForeColor.RGB = 220, 0, 0 }` |

**成员属性**

#### **Shapes.AddCallout**

在绘图画布上添加无框线的标注。

**语法**

**express.AddCallout**

*express*   一个代表 **Shapes** 对象的变量。

**说明**

使用 **AddShape** 方法可以插入多种不同的标注，例如气球和云朵。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在新创建的绘图画布上添加标注。*/ function NewCanvasCallout(){     //Add drawing canvas to the active document     let shpCanvas = ActiveDocument.Shapes.AddCanvas(150, 150, 200, 300)      //Add callout to the drawing canvas     shpCanvas.CanvasItems.AddCallout(msoCalloutTwo, 100, 40, 150, 75) }` |

#### **Shapes.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Shapes** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Shapes.Count**

返回一个 **Long** 类型的值，该值代表集合中图形的数量。只读。

**语法**

**express.Count**

*express*   一个代表 **Shapes** 对象的变量。

#### **Shapes.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Shapes** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Shapes.Parent**

返回一个 **Object** 类型值，该值代表指定 **Shapes** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Shapes** 对象的变量。

适用环境：web

适用平台：windows/linux