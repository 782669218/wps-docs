**Shapes**



指定的幻灯片上所有 [**Shape**](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape) 对象的集合。

**说明**

每个 **Shape** 对象代表绘图层中的一个对象，例如自选图形、任意多边形、OLE 对象或图片。

| ![img](https://qn.cache.wpscdn.cn/gif/close.gif) 注释        |
| ------------------------------------------------------------ |
| 若要使用文档中的形状的子集（例如，只对文档中的自选图形或选定的形状进行操作），则必须构造一个包含要使用的形状的 [**ShapeRange**](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange) 集合。有关如何使用一个形状或同时使用多个形状的概述，请参阅 使用形状（图形对象）。 |

## 示例

使用 **Shapes** 属性返回 **Shapes** 集合。以下示例选择当前演示文稿中的所有形状。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActivePresentation.Slides.Item(1).Shapes.SelectAll()` |

| ![img](https://qn.cache.wpscdn.cn/gif/close.gif) 注释        |
| ------------------------------------------------------------ |
| 若要同时对文档中的所有形状进行某种操作（例如删除或设置一个属性），可不带参数使用 [**Range**](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Shapes/Shapes%20.htm#Shapes.Range) 方法创建一个 **ShapeRange** 对象，该对象包含 **Shapes** 集合中的所有形状，然后对 **ShapeRange** 对象应用相应的属性或方法。 |

使用 

**AddCallout**

 、

**AddComment**

 、

**AddConnector**

 、

**AddCurve**

 、

**AddLabel**

 、

**AddLine**

 、

**AddMediaObject**

 、

**AddOLEObject**

 、

**AddPicture**

 、

**AddPlaceholder**

 、

**AddPolyline**

 、

**AddShape**

 、

**AddTable**

 、

**AddTextbox**

 、

**AddTextEffect**

 或 

**AddTitle**

 方法创建一个新形状并将其添加到 

Shapes

 集合中。联合使用 **BuildFreeform**

 方法和 **ConvertToShape**

 方法创建一个新任意多边形并将其添加到集合中。以下示例在活动演示文稿中添加一个矩形。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActivePresentation.Slides.Item(1).Shapes.AddShape(msoShapeRectangle, 50, 50, 100, 200)` |

使用 **Shapes**(*index*) 返回单个 **Shape** 对象，其中 *index* 是形状的名称或索引号。

使用 

Shapes.Range

(

index

) 返回一个 

**ShapeRange**

 集合，该集合代表 

Shapes

 集合的一个子集，其中 

index

 是形状的名称或索引号，或由形状的名称或索引号组成的数组。以下示例设置活动演示文稿中第一个和第三个形状的填充图案。

 

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActivePresentation.Slides.Item(1).Shapes.Range([1, 3]).Fill .Patterned(msoPatternHorizontalBrick)` |

使用 **Shapes.Placeholders**(*index*) 返回一个代表占位符的 **Shape** 对象，其中 *index* 是占位符编号。如果指定的幻灯片有标题，则使用 **Shapes.Placeholders(1)** 或 **Shapes.Title** 返回标题占位符。以下示例在活动演示文稿中添加一张幻灯片并为标题和副标题添加文本（副标题是此版式的幻灯片中的第二个占位符）。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myShape = Application.ActivePresentation.Slides.Add(1, ppLayoutTitle).Shapes     myShape.Title.TextFrame.TextRange.Text = "This is the title text"     myShape.Placeholders.Item(2).TextFrame.TextRange.Text = "This is subtitle text"  } ` |

**方法**

|                                                              | 名称               | 说明                                                         |
| ------------------------------------------------------------ | ------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddChart**       | 向演示文稿中添加图表。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddConnector**   | 创建一个连接符。返回一个代表新连接符的 **Shape** 对象。添加连接符时，它没有连接任何对象。使用 **BeginConnect** 和 **EndConnect** 方法可将连接符的头和尾连接到文档中的其他形状上。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddLine**        | 创建线条。返回一个代表新线条的 **Shape** 对象。              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddMediaObject** | 创建多媒体对象。返回一个代表新多媒体对象的 [Shape ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddOLEObject**   | 创建 OLE 对象。返回一个代表新 OLE 对象的 **Shape** 对象。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddPicture**     | 从现有文件创建图片。返回一个代表新图片的 **Shape** 对象。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddPlaceholder** | 恢复幻灯片上以前已删除的占位符。返回一个 **Shape** 代表已恢复占位符的对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddShape**       | 创建一个自选形状。返回一个代表新自选形状的 **Shape** 对象。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddTable**       | 将表格形状添加到幻灯片中。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddTextEffect**  | 创建艺术字对象。返回一个代表新艺术字对象的 **Shape** 对象。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddTextbox**     | 创建一个文本框。返回一个代表新文本框的 **Shape** 对象。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**           | 从指定集合中返回单个对象。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Paste**          | 在 z 顺序的最上端将剪贴板上的形状、幻灯片或文本粘贴到指定的 **Shapes** 集合中。粘贴的每个对象都会成为指定的 **Shapes** 集合的成员。如果剪贴板包含全部幻灯片，则这些幻灯片将作为包含幻灯片图像的形状粘贴。如果剪贴板包含文本范围，则该文本将被粘贴到一个新创建的 **TextFrame** 形状中。返回一个代表粘贴对象的 **ShapeRange** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Range**          | 返回一个代表 **Shapes** 集合中的形状子集的 **ShapeRange** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectAll**      | 选定所有形状（**Shapes** 集合中）或所有图表节点（ **DiagramNodes** 或 **DiagramNodeChildren** 集合中）。 |

**属性**

|                                                              | 名称             | 说明                                                         |
| ------------------------------------------------------------ | ---------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**  | 返回一个 **Application** 对象，该对象表示指定对象的创建者。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**        | 返回指定集合中的对象数目。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasTitle**     | 返回指定幻灯片上的对象集合是否包含标题占位符。只读 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Placeholders** | 返回一个 **Placeholders** 集合，该集合代表幻灯片上所有占位符的集合。集合中的每个占位符可包含文本、图表、表格、组织机构图或其他对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Title**        | 返回一个代表幻灯片标题的 **Shape** 对象。只读。              |

**成员方法**

#### **Shapes.AddChart**

向演示文稿中添加图表。

**语法**

**express.AddChart(Type, Left, Top, Width, Height)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**     |
| -------- | ------------- | ------------ | ------------ |
| *Type*   | 可选          | **Variant**  | 图表的类型。 |
| *Left*   | 可选          | **Number**   | 左侧尺寸。   |
| *Top*    | 可选          | **Number**   | 顶端尺寸。   |
| *Width*  | 可选          | **Number**   | 图表的宽度。 |
| *Height* | 可选          | **Number**   | 图表的高度。 |

#### **Shapes.AddConnector**

创建一个连接符。返回一个代表新连接符的 **Shape** 对象。添加连接符时，它没有连接任何对象。使用 **BeginConnect** 和 **EndConnect** 方法可将连接符的头和尾连接到文档中的其他形状上。

**语法**

**express.AddConnector(Type, BeginX, BeginY, EndX, EndY)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                 |
| -------- | ------------- | ------------ | -------------------------------------------------------- |
| *Type*   | 必选          | **Number**   | 连接符的类型。                                           |
| *BeginX* | 必选          | **Number**   | 连接符的起点相对于幻灯片左边缘的水平位置（以磅为单位）。 |
| *BeginY* | 必选          | **Number**   | 连接符的起点相对于幻灯片上边缘的垂直位置（以磅为单位）。 |
| *EndX*   | 必选          | **Number**   | 连接符的终点相对于幻灯片左边缘的水平位置（以磅为单位）。 |
| *EndY*   | 必选          | **Number**   | 连接符的终点相对于幻灯片上边缘的垂直位置（以磅为单位）。 |

**说明**

将一个连接符连接到某个形状时，如果必要，该连接符的长度和位置会自动调整。因此，如果要将一个连接符连接到其他形状，则与添加该连接符时指定的位置和长度无关。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加两个矩形，然后用曲线连接符将它们连接起来。*/ /*请注意，将连接符连接到矩形上时，连接符的长度和位置会自动调整；因此，它与添加标注时指定的位置和长度是无关的（长度不能为零）。*/ function test()  {     let shpShapes = Application.ActivePresentation.Slides.Item(1).Shapes     let shpFirst = shpShapes.AddShape(msoShapeRectangle, 100, 50, 200, 100)     let shpSecond = shpShapes.AddShape(msoShapeRectangle, 300, 300, 200, 100)     let myFormat = shpShapes.AddConnector(msoConnectorCurve, 0, 0, 100, 100).ConnectorFormat     myFormat.BeginConnect(shpFirst, 1)     myFormat.EndConnect(shpSecond, 1)     myFormat.Parent.RerouteConnections() }` |

#### **Shapes.AddLine**

创建线条。返回一个代表新线条的 **Shape** 对象。

**语法**

**express.AddLine(BeginX, BeginY, EndX, EndY)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                             |
| -------- | ------------- | ------------ | ---------------------------------------------------- |
| *BeginX* | 必选          | **Number**   | 线条起点相对于幻灯片左边缘的水平位置（以磅为单位）。 |
| *BeginY* | 必选          | **Number**   | 线条起点相对于幻灯片上边缘的垂直位置（以磅为单位）。 |
| *EndX*   | 必选          | **Number**   | 线条终点相对于幻灯片左边缘的水平位置（以磅为单位）。 |
| *EndY*   | 必选          | **Number**   | 线条终点相对于幻灯片上边缘的垂直位置（以磅为单位）。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例向 myDocument 中添加一条蓝色虚线。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     let myLine = myDocument.Shapes.AddLine(10, 10, 250, 250).Line     myLine.DashStyle = msoLineDashDotDot     myLine.ForeColor.RGB = (50, 0, 128) }` |

#### **Shapes.AddMediaObject**

创建多媒体对象。返回一个代表新多媒体对象的 [Shape ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象。

**语法**

**express.AddMediaObject(FileName, Left, Top, Width, Height)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------- | ------------- | ------------ | ------------------------------------------------------------ |
| *FileName* | 必选          | **String**   | 必选。String 类型。从中创建媒体对象的文件。如果未指定路径，则使用当前工作文件夹。 |
| *Left*     | 可选          | **Number**   | 媒体对象边界框左上角相对于文档左上角的位置（以磅为单位）。   |
| *Top*      | 可选          | **Number**   | 媒体对象边界框左上角相对于文档左上角的位置（以磅为单位）。   |
| *Width*    | 可选          | **Number**   | 媒体对象边界框的宽度（以磅为单位）。                         |
| *Height*   | 可选          | **Number**   | 媒体对象边界框的高度（以磅为单位）。                         |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加名为“Clock.avi”的影片。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     myDocument.Shapes.AddMediaObject("C:\\WINNT\\clock.avi", 5, 5, 100, 100) }` |

#### **Shapes.AddOLEObject**

创建 OLE 对象。返回一个代表新 OLE 对象的 **Shape** 对象。

**语法**

**express.AddOLEObject(Left, Top, Width, Height, ClassName, FileName, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Link)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型**    | **说明**                                                     |
| --------------- | ------------- | --------------- | ------------------------------------------------------------ |
| *Left*          | 可选          | **Number**      | 新对象左上角相对于幻灯片左上角的位置（以磅为单位）。默认值为 0（零）。 |
| *Top*           | 可选          | **Number**      | 新对象左上角相对于幻灯片左上角的位置（以磅为单位）。默认值为 0（零）。 |
| *Width*         | 可选          | **Number**      | OLE 对象的初始宽度（以磅为单位）。                           |
| *Height*        | 可选          | **Number**      | OLE 对象的初始高度（以磅为单位）。                           |
| *ClassName*     | 可选          | **String**      | OLE 长类名或待创建对象的 ProgID。必须为该对象指定 ClassName 或 FileName 参数，但不能同时指定两者。 |
| *FileName*      | 可选          | **String**      | 要创建的对象的源文件。如果未指定路径，则使用当前工作文件夹。必须为该对象指定 ClassName 或 FileName 参数，但不能同时指定两者。 |
| *DisplayAsIcon* | 可选          | **MsoTriState** | 决定是否将 OLE 对象显示为图标。                              |
| *IconFileName*  | 可选          | **String**      | 包含将要显示的图标的文件。                                   |
| *IconIndex*     | 可选          | **Number**      | IconFileName 中的图标索引。文件中第一个图标的索引号是 0（零）。如果 IconFileName 中不存在已知索引号的图标，则使用索引号为 1 的图标（文件中的第二个图标）。默认值为 0（零）。 |
| *IconLabel*     | 可选          | **String**      | 显示在图标下面的标签（标题）。                               |
| *Link*          | 可选          | **MsoTriState** | 决定是否将 OLE 对象链接到从中创建该对象的文件。如果已指定 ClassName 的值，则此参数必须是 msoFalse。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加命令按钮。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     myDocument.Shapes.AddOLEObject(100, 100, 150, 50, "Forms.CommandButton.1") }` |

#### **Shapes.AddPicture**

从现有文件创建图片。返回一个代表新图片的 **Shape** 对象。

**语法**

**express.AddPicture(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称**           | **必选/可选** | **数据类型**    | **说明**                                                     |
| ------------------ | ------------- | --------------- | ------------------------------------------------------------ |
| *FileName*         | 必选          | **String**      | 要在其中创建 OLE 对象的文件                                  |
| *LinkToFile*       | 必选          | **MsoTriState** | 确定是否将图片链接到从中创建该图片的文件。                   |
| *SaveWithDocument* | 必选          | **MsoTriState** | 确定是否将已链接的图片与其插入到的文档一起保存。如果 LinkToFile 为 msoFalse，则此参数必须为 msoTrue。 |
| *Left*             | 必选          | **Number**      | 图片左边缘相对于幻灯片左边缘的位置（以磅为单位）。           |
| *Top*              | 必选          | **Number**      | 图片上边缘相对于幻灯片上边缘的位置（以磅为单位）。           |
| *Width*            | 可选          | **Number**      | 图片的宽度（以磅为单位）。                                   |
| *Height*           | 可选          | **Number**      | 图片的高度（以磅为单位）。                                   |

#### **Shapes.AddPlaceholder**

恢复幻灯片上以前已删除的占位符。返回一个 **Shape** 代表已恢复占位符的对象。

**语法**

**express.AddPlaceholder(Type, Left, Top, Width, Height)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**          | **说明**                                                     |
| -------- | ------------- | --------------------- | ------------------------------------------------------------ |
| *Type*   | 必选          | **PpPlaceholderType** | 占位符的类型。只有在版式类型为 ppLayoutVerticalText、ppLayoutClipArtAndVerticalText、ppLayoutVerticalTitleAndText 或 ppLayoutVerticalTitleAndTextOverChart 的幻灯片中才能找到类型为 ppPlaceholderVerticalBody 或 ppPlaceholderVerticalTitle 的占位符。不能在用户界面中使用上述任何版式创建幻灯片；必须通过使用 Add 方法或通过设置现有幻灯片的 Layout 属性，以编程方式创建它们。 |
| *Left*   | 可选          | **Number**            | 占位符左上角相对于文档左上角的位置（以磅为单位）。           |
| *Top*    | 可选          | **Number**            | 占位符左上角相对于文档左上角的位置（以磅为单位）。           |
| *Width*  | 可选          | **Number**            | 占位符的宽度（以磅为单位）。                                 |
| *Height* | 可选          | **Number**            | 占位符的高度（以磅为单位）。                                 |

**说明**

如果从幻灯片中删除了多个指定类型的占位符，**AddPlaceholder** 方法将它们依次放回到幻灯片中，从原来索引号最小的占位符开始。

#### **Shapes.AddShape**

创建一个自选形状。返回一个代表新自选形状的 **Shape** 对象。

**语法**

**express.AddShape(Type, Left, Top, Width, Height)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**         | **说明**                                               |
| -------- | ------------- | -------------------- | ------------------------------------------------------ |
| *Type*   | 必选          | **MsoAutoShapeType** | 指定要创建的自选形状的类型。                           |
| *Left*   | 必选          | **Number**           | 自选形状左边缘相对于幻灯片左边缘的位置（以磅为单位）。 |
| *Top*    | 必选          | **Number**           | 自选形状上边缘相对于幻灯片上边缘的位置（以磅为单位）。 |
| *Width*  | 必选          | **Number**           | 自选形状的宽度（以磅为单位）。                         |
| *Height* | 必选          | **Number**           | 自选形状的高度（以磅为单位）。                         |

**说明**

若要更改所添加的自选形状的类型，可设置其 **AutoShapeType** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加矩形。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     myDocument.Shapes.AddShape(msoShapeRectangle, 50, 50, 100, 200) }` |

#### **Shapes.AddTable**

将表格形状添加到幻灯片中。

**语法**

**express.AddTable(NumRows, NumColumns, Left, Top, Width, Height)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型** | **说明**                                             |
| ------------ | ------------- | ------------ | ---------------------------------------------------- |
| *NumRows*    | 必选          | **Number**   | 表格的行数。                                         |
| *NumColumns* | 必选          | **Number**   | 表格的列数。                                         |
| *Left*       | 可选          | **Number**   | 从幻灯片左边缘到表格左边缘之间的距离（以磅为单位）。 |
| *Top*        | 可选          | **Number**   | 从幻灯片上边缘到表格上边缘之间的距离（以磅为单位）。 |
| *Width*      | 可选          | **Number**   | 新表格的宽度（以磅为单位）。                         |
| *Height*     | 可选          | **Number**   | 新表格的高度（以磅为单位）。                         |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在当前演示文稿的第二张幻灯片上创建一张新表格。*/ /*该表格具有三个行，四个列。它到幻灯片左、上边缘的距离均为 10 磅。*/ /*该新表格的宽度为 288 磅，这样，表格中四个列的宽度都是一英寸（一英寸等于 72 磅）。表格的高度设为 216 磅，这样，表格中三个行的高度都是一英寸。*/  Application.ActivePresentation.Slides.Item(2).Shapes .AddTable(3, 4, 10, 10, 288, 216)` |

#### **Shapes.AddTextEffect**

创建艺术字对象。返回一个代表新艺术字对象的 **Shape** 对象。

**语法**

**express.AddTextEffect(PresetTextEffect, Text, FontName, FontSize, FontBold, FontItalic, Left, Top)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称**           | **必选/可选** | **数据类型**            | **说明**                                                   |
| ------------------ | ------------- | ----------------------- | ---------------------------------------------------------- |
| *PresetTextEffect* | 必选          | **MsoPresetTextEffect** | 预置文字效果。                                             |
| *Text*             | 必选          | **String**              | 艺术字中的文字。                                           |
| *FontName*         | 必选          | **String**              | 艺术字中所用字体的名称。                                   |
| *FontSize*         | 必选          | **Number**              | 艺术字中所用字体的大小（以磅为单位）。                     |
| *FontBold*         | 必选          | **MsoTriState**         | 确定艺术字中使用的字体是否设为粗体。                       |
| *FontItalic*       | 必选          | **MsoTriState**         | 确定艺术字中使用的字体是否设为斜体。                       |
| *Left*             | 必选          | **Number**              | 艺术字边界框左边缘相对于幻灯片左边缘的位置（以磅为单位）。 |
| *Top*              | 必选          | **Number**              | 艺术字边界框上边缘相对于幻灯片上边缘的位置（以磅为单位）。 |

**说明**

向文档添加艺术字对象时，该对象的高度和宽度将自动根据所指定的文字的大小和数量来设置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加包含文本“Test”的艺术字。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     let newWordArt = myDocument.Shapes.AddTextEffect(msoTextEffect1, "Test", "Arial Black", 36, msoFalse, msoFalse, 10, 10) }` |

#### **Shapes.AddTextbox**

创建一个文本框。返回一个代表新文本框的 **Shape** 对象。

**语法**

**express.AddTextbox(Orientation, Left, Top, Width, Height)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型**           | **说明**                                                     |
| ------------- | ------------- | ---------------------- | ------------------------------------------------------------ |
| *Orientation* | 必选          | **MsoTextOrientation** | 文本方向。这些常量中的某些可能不可用，这取决于选择或安装的语言支持（例如，中国中文）。 |
| *Left*        | 必选          | **Nubmer**             | 文本框左边缘相对于幻灯片左边缘的位置（以磅为单位）。         |
| *Top*         | 必选          | **Nubmer**             | 文本框上边缘相对于幻灯片上边缘的位置（以磅为单位）。         |
| *Width*       | 必选          | **Nubmer**             | 文本框的宽度（以磅为单位）。                                 |
| *Height*      | 必选          | **Nubmer**             | 文本框的高度（以磅为单位）。                                 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加包含文本“Test Box”的文本框。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     myDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 200, 50).TextFrame.TextRange.Text = "Test Box" }` |

#### **Shapes.Item**

从指定集合中返回单个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                               |
| -------- | ------------- | ------------ | -------------------------------------- |
| *Index*  | 必选          | **Variant**  | 集合中要返回的单个对象的名称或索引号。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动演示文稿第一张幻灯片中名为“Rectangle 1”的形状的前景色设置为红色。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item("rectangle 1").Fill.ForeColor.RGB = (128, 0, 0)` |

#### **Shapes.Paste**

在 z 顺序的最上端将剪贴板上的形状、幻灯片或文本粘贴到指定的 **Shapes** 集合中。粘贴的每个对象都会成为指定的 **Shapes** 集合的成员。如果剪贴板包含全部幻灯片，则这些幻灯片将作为包含幻灯片图像的形状粘贴。如果剪贴板包含文本范围，则该文本将被粘贴到一个新创建的 **TextFrame** 形状中。返回一个代表粘贴对象的 **ShapeRange** 对象。

**语法**

**express.Paste()**

*express*   一个代表 **Shapes** 对象的变量。

**说明**

将剪贴板内容粘贴到窗口中之前，请使用 **ViewType** 属性设置窗口视图。下表显示了可以粘贴到每个视图中的内容。

| 视图                   | 可以从剪贴板中粘贴以下内容                                   |
| ---------------------- | ------------------------------------------------------------ |
| 幻灯片视图或备注页视图 | 形状、文本或整张幻灯片。如果从剪贴板粘贴一个幻灯片，该幻灯片的图像将作为嵌入对象被插入到幻灯片、母版或备注页中；如果选中形状，粘贴的文本将附加到形状文本之后；如果选中文本，粘贴的文本将替换选中的文本；如果选中任何其他对象，粘贴的文本将被放到它自己的文本框中。粘贴的形状将被放到 z 顺序的最上端且不会替换选中的形状。 |
| 大纲视图               | 文本或整张幻灯片。不能向大纲视图粘贴形状。粘贴的幻灯片将被插到光标所在的幻灯片之前。 |
| 幻灯片浏览视图         | 整张幻灯片。不能向幻灯片浏览视图粘贴形状或文本。粘贴的幻灯片将被插到光标处或演示文稿中最后选中的一张幻灯片之后。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动演示文稿第一张幻灯片的第一个形状复制到剪贴板中，然后将其粘贴到第二张幻灯片中。*/ function test()  {     let myPre = Application.ActivePresentation     myPre.Slides.Item(1).Shapes.Item(1).Copy()     myPre.Slides.Item(2).Shapes.Paste() }` |

#### **Shapes.Range**

返回一个代表 **Shapes** 集合中的形状子集的 **ShapeRange** 对象。

**语法**

**express.Range(Indexx)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Indexx* | 可选          | **Variant**  | 要包含在范围中的各个形状。可以是指定形状的索引号的 Integer、指定形状名称的 String，或者是包含整数或字符串的数组。如果省略此参数，则 Range 方法将返回指定集合中的所有对象。 |

**说明**

虽然可以使用 **Range** 方法返回任意数目的形状或幻灯片，但如果只需要返回该集合的一个成员，则使用 **Item** 方法会更为简单。例如，Shapes(1) 比 Shapes.Range(1) 简单，Slides(2) 比 Slides.Range(2) 简单。

若要为 **Index** 指定一个整数或字符串数组，可以使用 **Array** 函数。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 myDocument 中第一个形状和第三个形状的填充图案。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     myDocument.Shapes.Range([1, 3]).Fill.Patterned(msoPatternHorizontalBrick) }  /*本示例设置第一张幻灯片中名为 Oval 4 和 Rectangle 5 的形状的填充图案。*/ function test()  {     let myArray = ["Oval 4", "Rectangle 5"]     let myRange = Application.ActivePresentation.Slides.Item(1).Shapes.Range(myArray)     myRange.Fill.Patterned(msoPatternHorizontalBrick) }  /*本示例设置第一张幻灯片中所有形状的填充图案。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Range().Fill.Patterned(msoPatternHorizontalBrick)  /*本示例设置第一张幻灯片中第一个形状的填充图案。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     let myRange = myDocument.Shapes.Range(1)     myRange.Fill.Patterned(msoPatternHorizontalBrick) }` |

#### **Shapes.SelectAll**

选定所有形状（**Shapes** 集合中）或所有图表节点（ **DiagramNodes** 或 **DiagramNodeChildren** 集合中）。

**语法**

**express.SelectAll()**

*express*   一个代表 **Shapes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例选择 myDocument 中的所有形状。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     myDocument.Shapes.SelectAll() }` |

**成员属性**

#### **Shapes.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **Shapes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/ function test(pptPres)  {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }  /*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/ function test() {     let myShape = Application.ActivePresentation.Slides.Item(1).Shapes     for(let shpOle = 1; shpOle <= myShape.Count; shpOle++)      {         if(myShape.Item(shpOle).Type == msoLinkedOLEObject)         {             alert(myShape.Item(shpOle).OLEFormat.Application.Name)         }     } }` |

#### **Shapes.Count**

返回指定集合中的对象数目。

**语法**

**express.Count**

*express*   一个代表 **Shapes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭除当前窗口外的所有窗口。*/ function test() {     let myWindow = Application.Windows     for(let i = 2; i <= myWindow.Count; i++)      {         myWindow.Item(2).Close()     } }` |

#### **Shapes.HasTitle**

返回指定幻灯片上的对象集合是否包含标题占位符。只读 **MsoTriState** 类型。

**语法**

**express.HasTitle**

*express*   一个代表 **Shapes** 对象的变量。

**说明**

为ture，表示定的幻灯片上的对象集合包含标题占位符。

#### **Shapes.Placeholders**

返回一个 **Placeholders** 集合，该集合代表幻灯片上所有占位符的集合。集合中的每个占位符可包含文本、图表、表格、组织机构图或其他对象。只读。

**语法**

**express.Placeholders**

*express*   一个代表 **Shapes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在当前演示文稿中添加一张幻灯片，然后为该幻灯片标题（幻灯片的第一个占位符）和副标题添加文本。*/  function test() {     let myDocument = Application.ActivePresentation.Slides.Item(1)     let mySlide =  Application.ActivePresentation.Slides .Add(1, ppLayoutTitle).Shapes.Placeholders     mySlide.Item(1).TextFrame.TextRange.Text = "This is the title text"     mySlide.Item(2).TextFrame.TextRange.Text = "This is subtitle text" }` |

#### **Shapes.Title**

返回一个代表幻灯片标题的 **Shape** 对象。只读。

**语法**

**express.Title**

*express*   一个代表 **Shapes** 对象的变量。

**说明**

也可以使用 [**Shapes**](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Shapes/Shapes%20.htm#jsObject_Shapes) 或 **Placeholders** 集合的 **Item** 方法返回幻灯片标题。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 myDocument 上的标题文本。*/ function test() {     let myDocument = Application.ActivePresentation.Slides.Item(1)     myDocument.Shapes.Title.TextFrame.TextRange.Text = "Welcome!" }` |

适用环境：web

适用平台：windows/linux